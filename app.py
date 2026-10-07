import sys

# The Windows console defaults to a legacy code page that can't encode most
# non-Latin scripts, so any print() of a Japanese/Chinese/Arabic/etc. chat
# message (or an error quoting one) would raise and fail the whole request.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, 'reconfigure'):
        _stream.reconfigure(encoding='utf-8', errors='replace')

from flask import Flask, request, jsonify, send_from_directory, Response, session, redirect
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
import anthropic
import os
import json
import base64
import io
import re
import math
import random
import uuid
import secrets
import subprocess
import mimetypes
import platform
import tempfile
import shutil
import hashlib
import urllib.request
import urllib.error
import urllib.parse
import socket
import ipaddress
import threading
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
import docx as docx_lib
from docx.shared import RGBColor as DocxRGBColor, Pt as DocxPt
from docx.oxml.ns import qn as docx_qn
from docx.oxml import OxmlElement as DocxOxmlElement
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font as XlsxFont, PatternFill as XlsxFill, Alignment as XlsxAlignment
import pptx as pptx_lib
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from fpdf import FPDF
import csv as csv_lib
import email as email_lib
import email.policy
import html as html_lib
import extract_msg

load_dotenv()

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

# The one model every AI call in the app uses - change it here.
CLAUDE_MODEL = 'claude-sonnet-5-5'


# Where every *_FILE below actually lives. Defaults to the app's own
# directory (today's local-dev behavior, unchanged) - in production this is
# set to a mounted persistent volume (e.g. Railway), since the rest of the
# container's filesystem gets wiped on every redeploy/restart and this app's
# entire data layer is flat JSON files on disk, not a database.
DATA_DIR = os.getenv('DATA_DIR', '.')
os.makedirs(DATA_DIR, exist_ok=True)


def _data_path(filename):
    return os.path.join(DATA_DIR, filename)


# Every "now"/"today" in this app - business hours, the daily rollover, free-
# slot math, what Ashanti is told the current time is - has to be anchored to
# Francis's own timezone, not the server's. In production (Railway) that
# server clock is UTC, so calling the raw stdlib datetime.now()/date.today()
# used to read as several hours ahead of his actual wall clock: by
# mid-afternoon Eastern the server already thought business hours (9am-5pm)
# were over for the day, so nothing could be scheduled "today" even with
# hours of the real day left. Every such call in this file goes through
# these two instead. Single-user app, so one fixed zone for everything is
# correct - not derived per-request from the browser, which would have to be
# trusted and handled per to-do/event instead of being one global assumption.
APP_TIMEZONE = ZoneInfo('America/New_York')


def now_local():
    return datetime.now(APP_TIMEZONE).replace(tzinfo=None)


def today_local():
    return now_local().date()


# Signs the session cookie - generated once and persisted to .env so
# restarting the server doesn't silently log everyone out. Never checked
# into source control (this project isn't a git repo, but .env already holds
# the Anthropic/Outlook keys the same way).
def _get_or_create_flask_secret_key():
    existing = os.getenv('FLASK_SECRET_KEY')
    if existing:
        return existing
    key = secrets.token_hex(32)
    try:
        with open('.env', 'a', encoding='utf-8') as f:
            f.write(f'\nFLASK_SECRET_KEY={key}\n')
    except OSError:
        pass
    return key


app.secret_key = _get_or_create_flask_secret_key()


# --- Auth ------------------------------------------------------------------
# Real accounts (hashed passwords, Flask session cookie) so the data behind
# every other route in this file isn't wide open. Deliberately minimal for
# now - a single account (id 1, Francis) rather than open signup: "Get
# Started" on the landing page stays an inert placeholder (see
# showGetStartedPlaceholder in index.html) and Login's modal doubles as
# one-time account setup when users.json is still empty (see /auth/setup).
# Every other data store in this file stamps new records with userId (see
# current_user_id below) - always 1 today, but the column's there so real
# multi-user filtering is a small follow-up instead of a schema migration,
# whenever that's actually needed.
USERS_FILE = _data_path('users.json')
users_lock = threading.Lock()


def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []
    return []


def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, indent=2)


# Falls back to 1 (the sole account) even when nobody's logged in, e.g. for
# background/system-initiated writes - there's only ever one real owner of
# this data right now.
def current_user_id():
    return session.get('user_id') or 1


def _public_user(u):
    return {'id': u['id'], 'email': u['email'], 'name': u.get('name', '')}


@app.route('/auth/status')
def auth_status():
    users = load_users()
    user_id = session.get('user_id')
    user = next((u for u in users if u['id'] == user_id), None) if user_id else None
    return jsonify({
        'success': True,
        'loggedIn': bool(user),
        'setupNeeded': len(users) == 0,
        'user': _public_user(user) if user else None
    })


@app.route('/auth/setup', methods=['POST'])
def auth_setup():
    try:
        with users_lock:
            users = load_users()
            if users:
                return jsonify({'success': False, 'error': 'Setup already completed - use Login instead.'}), 400
            data = request.json or {}
            name = str(data.get('name', '')).strip()
            email = str(data.get('email', '')).strip().lower()
            password = data.get('password', '') or ''
            if not name or not email or not password:
                return jsonify({'success': False, 'error': 'Name, email, and password are all required.'}), 400
            if len(password) < 8:
                return jsonify({'success': False, 'error': 'Password must be at least 8 characters.'}), 400
            user = {
                'id': 1,
                'name': name,
                'email': email,
                'passwordHash': generate_password_hash(password),
                'createdAt': now_local().isoformat()
            }
            save_users([user])
        session.clear()
        # Not session.permanent - a plain (non-permanent) Flask session cookie
        # carries no expiry of its own, so the browser drops it the moment it
        # actually closes rather than keeping Francis logged in indefinitely.
        session['user_id'] = user['id']
        return jsonify({'success': True, 'user': _public_user(user)})
    except Exception as e:
        print(f"Auth setup error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/auth/login', methods=['POST'])
def auth_login():
    try:
        data = request.json or {}
        email = str(data.get('email', '')).strip().lower()
        password = data.get('password', '') or ''
        users = load_users()
        user = next((u for u in users if u['email'].lower() == email), None)
        if not user or not check_password_hash(user['passwordHash'], password):
            return jsonify({'success': False, 'error': 'Invalid email or password.'}), 401
        session.clear()
        # See the setup route for why this stays a non-permanent session.
        session['user_id'] = user['id']
        return jsonify({'success': True, 'user': _public_user(user)})
    except Exception as e:
        print(f"Auth login error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/auth/logout', methods=['POST'])
def auth_logout():
    session.clear()
    return jsonify({'success': True})


# Sent via Twilio SendGrid's HTTPS API directly (a plain urllib POST, same
# pattern as every other outbound call in this file - no reason to pull in
# the sendgrid package for one endpoint).
SENDGRID_API_KEY = os.getenv('SENDGRID_API_KEY')
SENDGRID_FROM_EMAIL = os.getenv('SENDGRID_FROM_EMAIL') or 'noreply@offload.com'
PASSWORD_RESET_TOKEN_TTL_MINUTES = 60


def _send_email_via_sendgrid(to_email, subject, text_body, html_body):
    if not SENDGRID_API_KEY:
        raise RuntimeError('SENDGRID_API_KEY is not set in .env - cannot send email.')
    payload = {
        'personalizations': [{'to': [{'email': to_email}]}],
        'from': {'email': SENDGRID_FROM_EMAIL},
        'subject': subject,
        'content': [
            {'type': 'text/plain', 'value': text_body},
            {'type': 'text/html', 'value': html_body}
        ]
    }
    req = urllib.request.Request(
        'https://api.sendgrid.com/v3/mail/send',
        data=json.dumps(payload).encode('utf-8'),
        headers={'Authorization': f'Bearer {SENDGRID_API_KEY}', 'Content-Type': 'application/json'},
        method='POST'
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        resp.read()


@app.route('/auth/forgot-password', methods=['POST'])
def auth_forgot_password():
    # Always the same response whether or not the email matched an account -
    # doesn't matter much with a single-user app, but it's the correct habit
    # (never lets a caller use this to discover which emails have accounts).
    generic_message = "If that email is registered, we've sent a reset link - check your inbox."
    try:
        data = request.json or {}
        email = str(data.get('email', '')).strip().lower()
        if not email:
            return jsonify({'success': False, 'error': 'Please enter your email.'}), 400

        token = None
        user = None
        with users_lock:
            users = load_users()
            user = next((u for u in users if u['email'].lower() == email), None)
            if user:
                # A hash of the token is stored, not the token itself - same
                # reasoning as passwordHash, so a leaked users.json can't be
                # used to reset the account even during the token's window.
                token = secrets.token_urlsafe(32)
                user['resetTokenHash'] = hashlib.sha256(token.encode('utf-8')).hexdigest()
                user['resetTokenExpiresAt'] = (
                    now_local() + timedelta(minutes=PASSWORD_RESET_TOKEN_TTL_MINUTES)
                ).isoformat()
                save_users(users)

        # The outbound HTTPS call to SendGrid runs unlocked, after the write
        # - same reasoning as every other network call in this file never
        # happening while a lock is held.
        if user:
            reset_url = f'{request.host_url.rstrip("/")}/?resetToken={token}'
            greeting = f"Hi {user['name']}," if user.get('name') else 'Hi,'
            try:
                _send_email_via_sendgrid(
                    user['email'],
                    'Reset your Offload password',
                    f"{greeting}\n\nSomeone (hopefully you) asked to reset your Offload password. "
                    f"This link expires in {PASSWORD_RESET_TOKEN_TTL_MINUTES} minutes:\n\n{reset_url}\n\n"
                    "If you didn't request this, you can safely ignore this email.",
                    f"<p>{greeting}</p>"
                    f"<p>Someone (hopefully you) asked to reset your Offload password. "
                    f"This link expires in {PASSWORD_RESET_TOKEN_TTL_MINUTES} minutes:</p>"
                    f'<p><a href="{reset_url}">{reset_url}</a></p>'
                    "<p>If you didn't request this, you can safely ignore this email.</p>"
                )
            except Exception as e:
                print(f"Password reset email send error: {e}")
                return jsonify({'success': False, 'error': 'Could not send the reset email - try again in a moment.'}), 500
        return jsonify({'success': True, 'message': generic_message})
    except Exception as e:
        print(f"Forgot password error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/auth/reset-password', methods=['POST'])
def auth_reset_password():
    try:
        data = request.json or {}
        token = str(data.get('token', '') or '')
        password = data.get('password', '') or ''
        if not token:
            return jsonify({'success': False, 'error': 'Missing or invalid reset link.'}), 400
        if len(password) < 8:
            return jsonify({'success': False, 'error': 'Password must be at least 8 characters.'}), 400

        token_hash = hashlib.sha256(token.encode('utf-8')).hexdigest()
        with users_lock:
            users = load_users()
            user = next((u for u in users if u.get('resetTokenHash') == token_hash), None)
            if not user:
                return jsonify({'success': False, 'error': 'This reset link is invalid or has already been used.'}), 400
            expires_at = user.get('resetTokenExpiresAt')
            if not expires_at or datetime.fromisoformat(expires_at) < now_local():
                return jsonify({'success': False, 'error': 'This reset link has expired - request a new one.'}), 400
            user['passwordHash'] = generate_password_hash(password)
            user.pop('resetTokenHash', None)
            user.pop('resetTokenExpiresAt', None)
            save_users(users)
        session.clear()
        session['user_id'] = user['id']
        return jsonify({'success': True, 'user': _public_user(user)})
    except Exception as e:
        print(f"Reset password error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500



_AUTH_EXEMPT_PATHS = {
    '/', '/auth/status', '/auth/setup', '/auth/login', '/auth/logout',
    '/auth/forgot-password', '/auth/reset-password'
}
_AUTH_EXEMPT_PREFIXES = ('/avatars/', '/backgrounds/')


@app.before_request
def _require_login():
    # All of the frontend's own API calls are relative (see index.html) -
    # they always resolve to whatever host actually served the page, so the
    # session cookie reaches them regardless of whether that's localhost,
    # 127.0.0.1, or the real deployed domain. No host canonicalizing needed.
    if request.method == 'OPTIONS':
        return
    if request.path in _AUTH_EXEMPT_PATHS or request.path.startswith(_AUTH_EXEMPT_PREFIXES):
        return
    if not session.get('user_id'):
        return jsonify({'success': False, 'error': 'Not authenticated'}), 401
    maybe_run_daily_kb_merge()


# --- Attachments -----------------------------------------------------------
# A single generic file store shared by every item type that can carry a
# file (To-Do, Calendar events, Discussion Topics, Tasks and Projects) -
# rather than one bespoke upload path per feature. Each upload gets its own
# folder (named by a random id) holding exactly one file under its original
# name, so serving it back needs no separate metadata store - the folder
# listing IS the metadata. The item itself (todo/event/topic/task) just
# holds a small {id, name, size, mimeType} descriptor pointing at one of
# these folders; Tasks aren't server-persisted at all (see workspaceTasks in
# index.html - they live in localStorage), so for those the descriptor is
# stored client-side while the actual bytes still live here.
ATTACHMENTS_DIR = _data_path('attachments')
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)
attachments_lock = threading.Lock()


def _attachment_file_path(attachment_id):
    folder = os.path.join(ATTACHMENTS_DIR, attachment_id)
    if not os.path.isdir(folder):
        return None
    names = os.listdir(folder)
    return os.path.join(folder, names[0]) if names else None


@app.route('/attachments', methods=['POST'])
def attachments_upload():
    try:
        file = request.files.get('file')
        if not file or not file.filename:
            return jsonify({'success': False, 'error': 'No file provided'}), 400
        attachment_id = uuid.uuid4().hex
        safe_name = os.path.basename(file.filename)
        with attachments_lock:
            folder = os.path.join(ATTACHMENTS_DIR, attachment_id)
            os.makedirs(folder, exist_ok=True)
            dest = os.path.join(folder, safe_name)
            file.save(dest)
            size = os.path.getsize(dest)
        mime_type = file.mimetype or mimetypes.guess_type(safe_name)[0] or 'application/octet-stream'
        return jsonify({
            'success': True,
            'attachment': {'id': attachment_id, 'name': safe_name, 'size': size, 'mimeType': mime_type}
        })
    except Exception as e:
        print(f"Attachment upload error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/attachments/<attachment_id>', methods=['GET'])
def attachments_get(attachment_id):
    path = _attachment_file_path(attachment_id)
    if not path:
        return ('Not found', 404)
    directory, filename = os.path.split(path)
    return send_from_directory(directory, filename)


@app.route('/attachments/<attachment_id>', methods=['DELETE'])
def attachments_delete(attachment_id):
    try:
        _delete_attachment_folder(attachment_id)
        return jsonify({'success': True})
    except Exception as e:
        print(f"Attachment delete error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


def _delete_attachment_folder(attachment_id):
    with attachments_lock:
        folder = os.path.join(ATTACHMENTS_DIR, attachment_id)
        if os.path.isdir(folder):
            shutil.rmtree(folder)


# Called wherever a whole item (to-do, calendar event, discussion topic) with
# attachments gets deleted outright, so its files don't linger as orphaned
# folders on disk - see the "takes up space" discussion with Francis.
def _delete_item_attachments(item):
    for att in (item.get('attachments') or []):
        att_id = att.get('id') if isinstance(att, dict) else None
        if att_id:
            _delete_attachment_folder(att_id)

# --- Claude usage log + prompt caching -------------------------------------
# Every call to Claude goes through claude_create so what it actually costs is
# on record (usage_log.jsonl, one line per call - see /usage/summary and
# Settings > Usage) instead of guessed at. The purpose of a call is simply the
# name of the function that made it.
USAGE_LOG_FILE = _data_path('usage_log.jsonl')
usage_log_lock = threading.Lock()

# USD per million tokens (cache_write is the 5-minute cache), and per web search.
MODEL_PRICING = {
    'claude-sonnet-5': {'input': 2.00, 'output': 10.00, 'cache_write': 2.50, 'cache_read': 0.20},
    'claude-sonnet-5-5': {'input': 2.00, 'output': 10.00, 'cache_write': 2.50, 'cache_read': 0.20},
    'claude-haiku-4-5-20251001': {'input': 1.00, 'output': 5.00, 'cache_write': 1.25, 'cache_read': 0.10},
}
WEB_SEARCH_COST_EACH = 0.01


def _estimate_call_cost(model, input_tokens, output_tokens, cache_write, cache_read, web_searches):
    price = MODEL_PRICING.get(model) or MODEL_PRICING['claude-sonnet-5']
    return (
        input_tokens * price['input'] + output_tokens * price['output']
        + cache_write * price['cache_write'] + cache_read * price['cache_read']
    ) / 1_000_000 + web_searches * WEB_SEARCH_COST_EACH


def _log_claude_usage(purpose, agent, model, response):
    try:
        usage = response.usage
        server_tool = getattr(usage, 'server_tool_use', None)
        web_searches = int(getattr(server_tool, 'web_search_requests', 0) or 0) if server_tool else 0
        input_tokens = int(getattr(usage, 'input_tokens', 0) or 0)
        output_tokens = int(getattr(usage, 'output_tokens', 0) or 0)
        cache_write = int(getattr(usage, 'cache_creation_input_tokens', 0) or 0)
        cache_read = int(getattr(usage, 'cache_read_input_tokens', 0) or 0)
        entry = {
            'ts': now_local().isoformat(), 'purpose': purpose, 'agent': agent, 'model': model,
            'input_tokens': input_tokens, 'output_tokens': output_tokens,
            'cache_write_tokens': cache_write, 'cache_read_tokens': cache_read,
            'web_searches': web_searches,
            'cost': round(_estimate_call_cost(model, input_tokens, output_tokens, cache_write, cache_read, web_searches), 6)
        }
        with usage_log_lock:
            with open(USAGE_LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(json.dumps(entry) + '\n')
    except Exception as e:
        print(f"Usage log error: {e}")


def claude_create(log_agent=None, log_purpose=None, **kwargs):
    purpose = log_purpose or sys._getframe(1).f_code.co_name
    response = client.messages.create(**kwargs)
    _log_claude_usage(purpose, log_agent, kwargs.get('model'), response)
    return response


# The system prompt goes to Claude as blocks so the parts that don't change
# between messages can be cached (read back at a tenth of the normal input
# price for ~5 minutes): the agent's personality + team knowledge, then the
# business notes. Anything that changes message to message (today's calendar,
# task context, an abbreviated older history) comes last, uncached - a change
# anywhere only invalidates what comes after it.
def build_system_blocks(static_text, notes_text, volatile_text):
    blocks = [{'type': 'text', 'text': static_text, 'cache_control': {'type': 'ephemeral'}}]
    if notes_text.strip():
        blocks.append({'type': 'text', 'text': notes_text, 'cache_control': {'type': 'ephemeral'}})
    if volatile_text.strip():
        blocks.append({'type': 'text', 'text': volatile_text})
    return blocks


KNOWLEDGE_BASE_FILE = _data_path('knowledge_base.json')
THOUGHT_STATUS_FILE = _data_path('thought_status.json')
PERSONAL_KNOWLEDGE_BASE_FILE = _data_path('personal_knowledge_base.json')
PERSONAL_THOUGHT_STATUS_FILE = _data_path('personal_thought_status.json')
KB_NOTES_FILE = _data_path('kb_notes.json')
KB_CONNECTIONS_FILE = _data_path('kb_connections.json')

ALL_AGENTS = ['manny', 'sasha', 'mark', 'kat', 'scott', 'tasha', 'techi', 'ashanti', 'lana']

# Static placeholder question banks for "Learn About Your Firm".
# One question is shown per agent per calendar day, cycling through the list.
QUESTION_BANKS = {
    'manny': [
        "What's the single biggest priority for the business this quarter?",
        "How many people (including you) currently work in the firm?",
        "What does a typical week look like for you as the owner?",
        "What's one decision you're currently sitting on?",
        "Where do you see the firm in 12 months?",
    ],
    'sasha': [
        "Which social platforms does the firm currently use, if any?",
        "Who is the ideal client you want to attract through social media?",
        "Do you have any brand guidelines (colors, tone, logo) I should know about?",
        "What's a recent win you'd want to show off publicly?",
        "How often would you like to post new content?",
    ],
    'mark': [
        "How do most of your clients currently find you?",
        "What's your average client value or engagement size?",
        "What's the biggest objection you hear from prospects?",
        "Do you have a formal sales process today, or is it ad hoc?",
        "What's your current client retention like year over year?",
    ],
    'kat': [
        "How would you describe the firm's voice — formal, friendly, technical?",
        "Is there any messaging or copy you're currently unhappy with?",
        "Who is reading your website and marketing materials most often?",
        "Do you have any phrases or words the brand should avoid?",
        "What's the one thing you want every client to understand about the firm?",
    ],
    'scott': [
        "How many people do you plan to hire in the next 6 months?",
        "What roles are hardest to fill right now?",
        "What does your onboarding process look like today?",
        "What traits matter most to you in a new hire?",
        "Do you have any team turnover concerns right now?",
    ],
    'tasha': [
        "How many clients do you currently serve?",
        "Do you specialize in a niche (crypto, real estate, small business, etc.)?",
        "What's your busiest season, and how do you handle the workload?",
        "What compliance or regulatory concerns keep you up at night?",
        "Roughly what's the firm's annual revenue range?",
    ],
    'techi': [
        "What software or tools does the firm rely on day to day?",
        "Is there a manual process you wish was automated?",
        "How do you currently store and secure client data?",
        "Have you had any tech outages or data issues recently?",
        "What's your budget appetite for new tools this year?",
    ],
    'ashanti': [
        "What's a task you keep putting off that I could help track?",
        "How do you currently manage your schedule and to-dos?",
        "What's one recurring deadline you don't want to miss?",
        "Do you prefer daily, weekly, or as-needed check-ins from me?",
        "What does a productive day look like for you?",
    ],
    'lana': [
        "Do any of your clients speak a language other than English at home?",
        "Which languages come up most in your client conversations or paperwork?",
        "Is there a language you've always wanted to learn, for work or for yourself?",
        "How much time could you realistically spare for language practice in a normal week?",
        "Do you ever translate documents or emails for clients today, and how?",
    ],
}

# Static placeholder idea banks for "Just A Thought".
# One idea is shown per agent per calendar day, cycling through the list.
# Phrased as tasks the agent itself will go do (first person, concrete, doable
# with the tools they actually have - mainly create_file), not general business
# advice for Francis to act on himself. Accepting one creates a real Workspace
# task assigned to that agent, which Start then actually runs.
IDEA_BANKS = {
    'manny': [
        "I'll put together a one-page team accountability tracker you can use for a weekly standup.",
        "I'll draft a quarterly goals review template so everyone's rowing in the same direction.",
        "I'll write up a first draft of your core processes as a doc, so they don't live only in your head.",
    ],
    'sasha': [
        "I'll draft a behind-the-scenes post about your process to help build trust with prospects.",
        "I'll put together a short client-testimonial request template you can send out.",
        "I'll build a one-month content calendar outline so posting doesn't depend on finding time each day.",
    ],
    'mark': [
        "I'll draft a simple referral incentive one-pager to turn happy clients into a lead source.",
        "I'll write a follow-up email template for prospects who went quiet.",
        "I'll put together a one-page case study template to help close new business faster.",
    ],
    'kat': [
        "I'll draft a few stronger headline options for your website that state who you help.",
        "I'll write a short FAQ page that pre-answers objections before they reach Mark.",
        "I'll put together a quick terminology/style sheet so your materials stay consistent.",
    ],
    'scott': [
        "I'll draft a structured interview scorecard to speed up hiring decisions.",
        "I'll put together a simple 30/60/90 day onboarding checklist for new hires.",
        "I'll write a short exit-interview template to help surface retention issues early.",
    ],
    'tasha': [
        "I'll put together a filing-deadline checklist doc so nothing sneaks up on you.",
        "I'll draft a client-facing document checklist to cut down on back-and-forth emails.",
        "I'll build a quick list of clients who might benefit from a mid-year check-in.",
    ],
    'techi': [
        "I'll write up a simple backup routine doc so nothing's lost if something breaks.",
        "I'll draft a shared password-manager rollout checklist for the team.",
        "I'll document one repetitive weekly task so it's ready to automate.",
    ],
    'ashanti': [
        "I'll put together a shared team calendar template to cut down on scheduling back-and-forth.",
        "I'll draft a task-batching guide to help save context-switching time.",
        "I'll build a weekly open-items review template so nothing slips through.",
    ],
    'lana': [
        "I'll draft a short bilingual welcome email for clients who'd be more comfortable in another language.",
        "I'll put together a one-page cheat sheet of key tax-season phrases in the language your clients speak most.",
        "I'll write up a 90-day starter plan for a language you'd like to learn, sized to the time you actually have.",
    ],
}


# Static placeholder question banks for "About You" - personal (not business)
# questions, each grounded in the specific agent's own interests/expertise so it
# feels like it's coming from that person rather than a generic intake form.
PERSONAL_QUESTION_BANKS = {
    'manny': [
        "Do you have any hobbies you make time for outside the practice, or does work take up most of your bandwidth?",
        "Are you into photography, music, or anything creative in your downtime?",
        "Do you follow any sports? I could always use someone to argue standings with.",
        "When's the last time you took a real day off, no work at all?",
        "Do you cook, or is takeout doing most of the heavy lifting these days?",
    ],
    'sasha': [
        "Do you keep up with any social media trends yourself, or is that strictly a work thing for you?",
        "Any shows, podcasts, or accounts you're low-key obsessed with right now?",
        "Do you have any plants or pets keeping you company at home?",
        "What's something you're into right now that has nothing to do with taxes?",
        "Do you ever thrift or shop vintage, or is that not your thing?",
    ],
    'mark': [
        "Do you golf, or is that just my thing I keep bringing up?",
        "Are you the networking-event type, or more heads-down?",
        "Got any home projects going right now?",
        "Do you ever mentor anyone outside of work, or is your plate too full for that right now?",
        "What's something you're personally working toward this year, outside the business?",
    ],
    'kat': [
        "Do you read much, or is your reading time eaten up by client emails these days?",
        "Any interest in theater, art, or museums, or is that not really your scene?",
        "Do you have any kind of movement practice — yoga, gym, walks — that helps you reset?",
        "How do you personally unwind after a long day with clients?",
        "Are you a nutrition-conscious person, or is coffee the main food group most days?",
    ],
    'scott': [
        "Do you make time to work out, or has that fallen off with everything going on?",
        "Any true crime shows or podcasts you've been into lately?",
        "Do you get out for drinks or dinner with friends much, or is it mostly work these days?",
        "Who's someone who mentored you early in your career?",
        "Do you have a go-to way to unwind — a hobby, a drink, a show — after a long week?",
    ],
    'tasha': [
        "Do you get outdoors much, or is it mostly desk time these days?",
        "Are you a puzzle person, or does your brain get enough of that at work?",
        "Any documentaries or deep-dive shows you've been into lately?",
        "Do you garden or grow anything, or is your green thumb strictly hypothetical?",
        "What's something you do that has absolutely nothing to do with taxes?",
    ],
    'techi': [
        "Are you much of a gadget person yourself, or strictly business tools?",
        "Any sci-fi shows, books, or movies you're into?",
        "Do you collect anything, or is that just a me thing?",
        "What's your relationship with your phone like — do you actually unplug, or always on?",
        "Do you geek out on anything outside of work? I promise not to over-explain it.",
    ],
    'ashanti': [
        "Do you journal or keep any kind of personal system for yourself, or is that just for the business?",
        "How are you sleeping and eating these days, actually?",
        "Do you listen to audiobooks or podcasts, or is your commute more of a decompress-in-silence thing?",
        "Who do you lean on when things get overwhelming — work or personal?",
        "What does a good, restful day look like for you, if you ever get one?",
    ],
    'lana': [
        "Do you speak, or have you ever studied, another language? Even a few leftover high-school phrases count.",
        "Is there a country or culture you've always wanted to visit?",
        "Do you listen to music in other languages, or is it mostly English?",
        "Do you have a dog, or a pet who's quietly running your household?",
        "What's a dish you love that comes from somewhere else? I'm always collecting food vocabulary.",
    ],
}

# Static placeholder idea banks for "Suggestions for You" - personal, not
# business, suggestions grounded in that agent's own interests/expertise. Like
# IDEA_BANKS, each is a task the agent itself will go do (first person, a
# document or researched list they can actually produce), never advice for
# Francis to carry out on his own - accepting one creates a real Workspace
# task assigned to that agent.
PERSONAL_IDEA_BANKS = {
    'manny': [
        "I'll draft a one-page plan for protecting one full day off each month, with a few ways to make it stick.",
        "I'll put together a short list of low-pressure chess and strategy games you can play in ten-minute breaks.",
        "I'll write up a simple phone-silent walk routine you can drop between meetings, with a few route ideas.",
    ],
    'sasha': [
        "I'll put together a short list of accounts and creators outside your industry that are worth following to unplug.",
        "I'll draft a simple evening no-phone routine you can try for a week, with a one-page tracker.",
        "I'll research a few low-maintenance houseplants that suit a busy schedule and write up the care basics.",
    ],
    'mark': [
        "I'll put together a short list of golf and mini-golf spots that make an easy half-day break.",
        "I'll draft a plan for one purely social event a month - a few ideas and a simple way to pick between them.",
        "I'll write up a list of quick 20-minute home projects you can knock out as a mental break from the practice.",
    ],
    'kat': [
        "I'll put together a short reading list of fiction picks for winding down before bed instead of scrolling.",
        "I'll draft a five-minute stretch and breathing routine for the end of a busy day.",
        "I'll write up three easy, screen-free dinner recipes with a simple plan for fitting one in each week.",
    ],
    'scott': [
        "I'll put together a list of 20-minute walk routes and a few podcast picks to pair with them.",
        "I'll draft a recurring friends-dinner plan, with a few ways to protect it on your calendar.",
        "I'll write up a short list of true crime podcasts to unwind with, with a one-line description of each.",
    ],
    'tasha': [
        "I'll research a few good short hikes and put together a one-page weekend guide.",
        "I'll put together a week of quick logic puzzles to start your mornings, with the answers.",
        "I'll write up a simple windowsill herb garden setup guide with a shopping list.",
    ],
    'techi': [
        "I'll draft a digital declutter checklist - unused apps, notifications to turn off - you can finish in 30 minutes.",
        "I'll put together a short list of sci-fi shows and retro games for a low-key escape.",
        "I'll write up a simple hard-stop routine for checking work on your phone at night, with the settings to change.",
    ],
    'ashanti': [
        "I'll put together a simple five-minute nightly brain-dump journal template you can print.",
        "I'll draft a plan for protecting one meal a day away from your desk, with easy meal ideas.",
        "I'll pull together a short list of audiobooks to unwind with, picked to fit your tastes.",
    ],
    'lana': [
        "I'll put together a short list of songs in another language, with a line-by-line meaning for one to start with.",
        "I'll draft a simple name-five-things-on-a-walk language game, with a vocabulary list to match.",
        "I'll write up a list of beginner children's books in a language you're curious about, and where to find them.",
    ],
}


def load_knowledge_base():
    if os.path.exists(KNOWLEDGE_BASE_FILE):
        try:
            with open(KNOWLEDGE_BASE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_knowledge_base(data):
    with open(KNOWLEDGE_BASE_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def today_str():
    return today_local().isoformat()


DAILY_ITEMS_PER_AGENT = 1


def get_today_questions(agent):
    bank = QUESTION_BANKS.get(agent, [])
    if not bank:
        return []
    base = today_local().toordinal()
    count = min(DAILY_ITEMS_PER_AGENT, len(bank))
    return [bank[(base + i) % len(bank)] for i in range(count)]


def get_today_ideas(agent):
    bank = IDEA_BANKS.get(agent, [])
    if not bank:
        return []
    base = today_local().toordinal()
    count = min(DAILY_ITEMS_PER_AGENT, len(bank))
    return [bank[(base + i) % len(bank)] for i in range(count)]


def load_thought_status():
    if os.path.exists(THOUGHT_STATUS_FILE):
        try:
            with open(THOUGHT_STATUS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_thought_status(data):
    with open(THOUGHT_STATUS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def load_personal_knowledge_base():
    if os.path.exists(PERSONAL_KNOWLEDGE_BASE_FILE):
        try:
            with open(PERSONAL_KNOWLEDGE_BASE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_personal_knowledge_base(data):
    with open(PERSONAL_KNOWLEDGE_BASE_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def get_today_personal_questions(agent):
    bank = PERSONAL_QUESTION_BANKS.get(agent, [])
    if not bank:
        return []
    base = today_local().toordinal()
    count = min(DAILY_ITEMS_PER_AGENT, len(bank))
    return [bank[(base + i) % len(bank)] for i in range(count)]


def get_today_personal_ideas(agent):
    bank = PERSONAL_IDEA_BANKS.get(agent, [])
    if not bank:
        return []
    base = today_local().toordinal()
    count = min(DAILY_ITEMS_PER_AGENT, len(bank))
    return [bank[(base + i) % len(bank)] for i in range(count)]


def load_personal_thought_status():
    if os.path.exists(PERSONAL_THOUGHT_STATUS_FILE):
        try:
            with open(PERSONAL_THOUGHT_STATUS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_personal_thought_status(data):
    with open(PERSONAL_THOUGHT_STATUS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


# --- Knowledge Base page (a single living, organized report per side) ---
# kb_notes.json holds exactly two strings - "firm" and "personal" - each the
# full current report shown on the Knowledge Base page. Nothing is ever stored
# as raw, unorganized text: a chat answer, typed input, or an uploaded file is
# folded into the existing report via merge_into_kb_report (Claude reconciles
# new info against what's already there); a direct manual edit on the page is
# cleaned up via organize_kb_text (Claude just re-organizes the replacement
# text, since it's already the whole intended content). Either way the page
# always reads as one organized summary, never a raw paragraph or a log.
def load_kb_notes():
    if os.path.exists(KB_NOTES_FILE):
        try:
            with open(KB_NOTES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_kb_notes(data):
    with open(KB_NOTES_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def kb_subject_label(side):
    return 'this tax practice / firm' if side == 'firm' else "Francis, the firm owner, personally (not business facts)"


# What chat sees of a side's connected files/folders. They form a searchable
# LIBRARY: the files' full text is NOT pasted into every message (a big
# spreadsheet re-sent on every message was the single biggest cost in the app)
# - instead this is a short, stable listing (name + headline overview), and the
# agent calls the search_library tool to pull the specific passages a question
# needs, only when it needs them. full=True is the old everything-included
# form, used by the explicit "ask the Knowledge Base" box (see /kb/ask), where
# the one question is the whole request.
def get_kb_connections_context(side, full=False):
    connections, _ = sync_all_kb_connections()
    if full:
        parts = []
        for conn in connections.get(side, []):
            if conn.get('missing'):
                continue
            content = (conn.get('raw_text') or '').strip() or (conn.get('summary') or '').strip()
            if not content:
                continue
            parts.append(f"From \"{conn['path']}\":\n{content[:100000]}")
        if not parts:
            return ""
        return (
            "\n\nLIVE CONNECTED FILES/FOLDERS (read directly from disk right now, not stored - "
            "if a connection is removed this section will simply stop appearing):\n" + "\n\n".join(parts)
        )

    lines = []
    for conn in connections.get(side, []):
        if conn.get('missing') or not (conn.get('raw_text') or '').strip():
            continue
        entry = f"- \"{conn['path']}\""
        summary = (conn.get('summary') or '').strip()
        if summary:
            entry += "\n  Overview (headline only - not for exact figures):\n" + "\n".join('  ' + l for l in summary.split('\n'))
        lines.append(entry)
    if not lines:
        return ""
    return (
        f"\n\nLIBRARY - connected files about {kb_subject_label(side)}, kept up to date. Their full contents are "
        "NOT included here; use the search_library tool for anything that depends on specifics from them "
        "(figures, dates, names, rows) rather than guessing or leaning on the overview:\n" + "\n".join(lines)
    )


# --- Library search ------------------------------------------------------
# Splits each connected file's stored text into chunks and finds the ones that
# match a query, so only a few hundred tokens of the right passages reach the
# model instead of the whole file. Plain keyword scoring (rarer words count
# for more) plus an exact-phrase and date-format bonus - no extra service.
LIBRARY_CHUNK_CHARS = 1800
LIBRARY_RESULT_CHARS = 7000
LIBRARY_SEARCH_MAX_RESULTS = 4
LIBRARY_STOPWORDS = {
    'the', 'a', 'an', 'and', 'or', 'of', 'to', 'in', 'on', 'for', 'is', 'was', 'what', 'whats', 'how', 'much',
    'many', 'did', 'do', 'does', 'we', 'our', 'my', 'me', 'i', 'it', 'at', 'by', 'with', 'from', 'this', 'that'
}


def library_has_content():
    connections = load_kb_connections()
    return any(
        not conn.get('missing') and (conn.get('raw_text') or '').strip()
        for side in ('firm', 'personal') for conn in connections.get(side, [])
    )


def _library_files():
    files = []
    connections = load_kb_connections()
    for side in ('firm', 'personal'):
        for conn in connections.get(side, []):
            raw = (conn.get('raw_text') or '').strip()
            if conn.get('missing') or not raw:
                continue
            # A folder's text is one "[Attached file: name]" block per file.
            pieces = re.split(r'(?m)^\[Attached file: (.+?)\]\s*$', raw)
            if len(pieces) >= 3:
                for i in range(1, len(pieces), 2):
                    files.append({'name': pieces[i], 'source': conn['path'], 'text': pieces[i + 1].strip()})
            else:
                files.append({'name': os.path.basename(conn['path'].rstrip('/\\')) or conn['path'],
                              'source': conn['path'], 'text': raw})
    return files


# Each chunk repeats its sheet's name and column row, so a chunk from the
# middle of a long daily-numbers sheet still says what each column is.
def _chunk_file_text(text):
    chunks = []
    header = ''
    current = []
    size = 0
    expecting_columns = False

    def flush():
        nonlocal current, size
        if current:
            chunks.append((header + '\n' if header else '') + '\n'.join(current))
        current, size = [], 0

    for line in text.split('\n'):
        if line.startswith('Sheet: '):
            flush()
            header = line
            expecting_columns = True
            continue
        if expecting_columns:
            header = header + '\nColumns: ' + line[:400]
            expecting_columns = False
            continue
        if current and size + len(line) > LIBRARY_CHUNK_CHARS:
            flush()
        current.append(line)
        size += len(line) + 1
    flush()
    return chunks or [text]


# A date in the query, written either way (2026-09-15 or 9/15/2026), also
# matches the other way spreadsheets commonly store it.
def _query_date_variants(query):
    variants = set()
    dates = [(int(y), int(m), int(d)) for y, m, d in re.findall(r'\b(\d{4})-(\d{1,2})-(\d{1,2})\b', query)]
    for m, d, y in re.findall(r'\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b', query):
        y = int(y)
        dates.append((y + 2000 if y < 100 else y, int(m), int(d)))
    for y, m, d in dates:
        variants.update({f'{y}-{m:02d}-{d:02d}', f'{m}/{d}/{y}', f'{m:02d}/{d:02d}/{y}', f'{m}/{d}/{y % 100:02d}'})
    return variants


def run_library_search(args):
    query = str(args.get('query') or '').strip()
    file_filter = str(args.get('file') or '').strip().lower()
    where = str(args.get('where') or 'match').strip().lower()
    part = args.get('part')

    files = _library_files()
    if file_filter:
        files = [f for f in files if file_filter in f['name'].lower() or file_filter in f['source'].lower()]
    if not files:
        return "The library has no readable files" + (" matching that name." if file_filter else " (nothing is connected, or the connection can't be read right now).")

    chunked = [(f['name'], _chunk_file_text(f['text'])) for f in files]

    def render(name, index, total, text):
        return f"--- {name} (part {index + 1} of {total}) ---\n{text}"

    results = []
    if part:
        try:
            wanted = int(part) - 1
        except (TypeError, ValueError):
            return 'part must be a number.'
        name, chunks = chunked[0]
        if not 0 <= wanted < len(chunks):
            return f'{name} has {len(chunks)} parts - ask for a part between 1 and {len(chunks)}.'
        results.append(render(name, wanted, len(chunks), chunks[wanted]))
    elif where in ('latest', 'start'):
        for name, chunks in chunked:
            picks = range(max(0, len(chunks) - 3), len(chunks)) if where == 'latest' else range(0, min(2, len(chunks)))
            for i in picks:
                results.append(render(name, i, len(chunks), chunks[i]))
    else:
        terms = [t for t in re.findall(r'[a-z0-9]+', query.lower()) if t not in LIBRARY_STOPWORDS and (len(t) > 1 or t.isdigit())]
        if not terms:
            return 'Give the specific words, names or dates to look for (as they would appear in the file), or use where="latest".'
        entries = [(name, i, len(chunks), chunk, chunk.lower()) for name, chunks in chunked for i, chunk in enumerate(chunks)]
        total = len(entries)
        patterns = {t: re.compile(r'(?<![a-z0-9])' + re.escape(t) + r'(?![a-z0-9])') for t in terms}
        doc_freq = {t: sum(1 for e in entries if patterns[t].search(e[4])) for t in terms}
        phrase = ' '.join(query.lower().split())
        date_variants = _query_date_variants(query)
        scored = []
        for name, i, n, chunk, lower in entries:
            score = 0.0
            for t in terms:
                count = len(patterns[t].findall(lower))
                if count:
                    score += math.log(1 + total / (1 + doc_freq[t])) * (1 + math.log(count))
            if score and len(phrase) > 3 and phrase in lower:
                score += 10
            if date_variants and any(v in lower for v in date_variants):
                score += 12
            if score > 0:
                scored.append((score, name, i, n, chunk))
        scored.sort(key=lambda x: -x[0])
        # Only passages close to the best match - a precise hit (an exact
        # date, a client's name) shouldn't drag in loosely related ones.
        best = scored[0][0] if scored else 0
        for score, name, i, n, chunk in [x for x in scored if x[0] >= best * 0.6][:LIBRARY_SEARCH_MAX_RESULTS]:
            results.append(render(name, i, n, chunk))
        if not results:
            return 'No passages matched. Try different words (as written in the file), a date like 2026-10-03, or where="latest" for the most recent entries.'

    output = '\n\n'.join(results)
    if len(output) > LIBRARY_RESULT_CHARS:
        output = output[:LIBRARY_RESULT_CHARS] + '\n...(cut off - ask for a specific part, or narrow the search)'
    return output


def merge_into_kb_report(side, new_info, source_label=None):
    notes = load_kb_notes()
    current_report = (notes.get(side) or '').strip()
    subject = kb_subject_label(side)
    source_note = f" (source: {source_label})" if source_label else ""

    if current_report:
        prompt = (
            f"You maintain a living knowledge-base report about {subject}. Here is the CURRENT report:\n\n"
            f"---\n{current_report}\n---\n\n"
            f"Here is NEW information just learned{source_note}:\n\n---\n{new_info}\n---\n\n"
            "Rewrite the report to incorporate this new information. Rules:\n"
            "- Keep every still-accurate existing fact - don't drop anything.\n"
            "- If the new information corrects or updates something already in the report, replace the old version "
            "with the corrected one instead of listing both.\n"
            "- Organize clearly with markdown headers (##), bullet points, and a table where tabular data genuinely "
            "fits (e.g. staff, figures by year) - make it easy to scan at a glance.\n"
            "- Plain factual third-person style - no meta-commentary like \"the user said\" or \"according to the file\".\n"
            "- Reply with ONLY the updated report in markdown - no preamble, no explanation, nothing else."
        )
    else:
        prompt = (
            f"You maintain a living knowledge-base report about {subject}. There's no report yet. "
            f"Here is the first piece of information{source_note}:\n\n---\n{new_info}\n---\n\n"
            "Write an initial report organizing this into clear markdown sections with headers (##) and bullet "
            "points (a table only if the information is genuinely tabular). Plain factual third-person style, no "
            "meta-commentary. Reply with ONLY the report in markdown - no preamble."
        )

    response = claude_create(
        model=CLAUDE_MODEL,
        max_tokens=1800,
        messages=[{'role': 'user', 'content': prompt}]
    )
    updated = "".join(
        block.text for block in response.content if getattr(block, 'type', None) == 'text'
    ).strip()

    notes[side] = updated
    save_kb_notes(notes)
    return updated


# --- Daily merge of check-in answers into the report ---------------------
# Answering a check-in question used to rewrite the entire report right away -
# an AI call that re-reads and re-writes the whole thing (about 2 cents each,
# growing with the report). The answer itself is already saved the moment it's
# submitted (agents see the latest 30 straight from there - see
# get_knowledge_base_context), so folding it into the report can wait: new
# answers are flagged merged=False, and all of a side's pending answers go into
# the report in ONE rewrite - at the first activity of each day, and whenever
# the Knowledge Base page is opened (so what's shown there is current). Entries
# saved before this existed have no flag and count as already merged.
KB_MERGE_STATE_FILE = _data_path('kb_merge_state.json')
kb_answers_lock = threading.Lock()
kb_merge_thread_lock = threading.Lock()
kb_merge_in_progress = False
_kb_daily_checked_date = None


def _kb_store(side):
    if side == 'firm':
        return load_knowledge_base, save_knowledge_base
    return load_personal_knowledge_base, save_personal_knowledge_base


def count_pending_kb_answers(side):
    load, _ = _kb_store(side)
    return sum(1 for entries in load().values() for e in entries if e.get('merged') is False)


def merge_pending_kb_answers(side):
    load, save = _kb_store(side)
    with kb_answers_lock:
        snapshot = [
            (agent, e.get('date', ''), e.get('question', ''), e.get('answer', ''))
            for agent, entries in load().items() for e in entries if e.get('merged') is False
        ]
    if not snapshot:
        return 0
    new_info = "\n\n".join(
        f"Q (asked by {agent.capitalize()}): {question}\nA (answered {date}): {answer}"
        for agent, date, question, answer in snapshot
    )
    # If this raises (network, API), nothing is marked merged and the answers
    # simply wait for the next run.
    merge_into_kb_report(side, new_info, source_label="answers to check-in questions")
    done = {(agent, question, answer) for agent, _, question, answer in snapshot}
    with kb_answers_lock:
        kb = load()
        for agent, entries in kb.items():
            for e in entries:
                if e.get('merged') is False and (agent, e.get('question', ''), e.get('answer', '')) in done:
                    e['merged'] = True
        save(kb)
    return len(snapshot)


# Runs the merge for both sides on a background thread (it's an AI call, so it
# never holds up the request that triggered it); a second trigger while one
# is running does nothing.
def kick_off_kb_merge():
    global kb_merge_in_progress
    with kb_merge_thread_lock:
        if kb_merge_in_progress:
            return False
        if not any(count_pending_kb_answers(side) for side in ('firm', 'personal')):
            return False
        kb_merge_in_progress = True

    def run():
        global kb_merge_in_progress
        try:
            for side in ('firm', 'personal'):
                try:
                    merge_pending_kb_answers(side)
                except Exception as e:
                    print(f"KB merge error ({side}): {e}")
        finally:
            with kb_merge_thread_lock:
                kb_merge_in_progress = False

    threading.Thread(target=run, daemon=True).start()
    return True


# "Once a day" without a scheduler: the first request of each new day claims
# the day (remembered in kb_merge_state.json so a restart doesn't repeat it)
# and kicks off the merge of everything answered before then.
def maybe_run_daily_kb_merge():
    global _kb_daily_checked_date
    today = today_str()
    if _kb_daily_checked_date == today:
        return
    _kb_daily_checked_date = today
    state = {}
    if os.path.exists(KB_MERGE_STATE_FILE):
        try:
            with open(KB_MERGE_STATE_FILE, 'r', encoding='utf-8') as f:
                state = json.load(f)
        except (json.JSONDecodeError, OSError):
            state = {}
    if state.get('lastDailyMerge') == today:
        return
    state['lastDailyMerge'] = today
    with open(KB_MERGE_STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f)
    kick_off_kb_merge()


# Used when Francis edits the report directly - the report should always read
# as an organized summary, even if what got typed/pasted into the textarea was
# a raw, unstructured paragraph. Unlike merge_into_kb_report, this has no
# separate "current report" to reconcile against - the edited text already IS
# the whole intended content, so it's cleaned up in place rather than merged
# with what came before.
def organize_kb_text(side, raw_text):
    raw_text = raw_text.strip()
    if not raw_text:
        return ''
    subject = kb_subject_label(side)
    prompt = (
        f"You maintain a living knowledge-base report about {subject}. Here is a raw draft to clean up:\n\n"
        f"---\n{raw_text}\n---\n\n"
        "Rewrite this as a clean, organized report: markdown headers (##), bullet points, and a table where "
        "tabular data genuinely fits. Keep every fact in the draft - don't drop or invent anything. Plain factual "
        "third-person style, no meta-commentary. Reply with ONLY the organized report in markdown - no preamble."
    )
    response = claude_create(
        model=CLAUDE_MODEL,
        max_tokens=1800,
        messages=[{'role': 'user', 'content': prompt}]
    )
    return "".join(
        block.text for block in response.content if getattr(block, 'type', None) == 'text'
    ).strip()


@app.route('/kb/data', methods=['GET'])
def kb_data():
    # Opening (or polling) the Knowledge Base page folds any waiting check-in
    # answers into the report in the background; the page's next refresh shows it.
    kick_off_kb_merge()
    connections, _ = sync_all_kb_connections()
    notes = load_kb_notes()
    return jsonify({
        'success': True,
        'firm': {'notes': notes.get('firm', '')},
        'personal': {'notes': notes.get('personal', '')},
        'connections': connections
    })


# Merges brand-new information into the existing report (used by the "add
# info" box, chat check-ins, and file uploads) - as opposed to /kb/notes below,
# which cleans up a full replacement text a direct edit already produced.
@app.route('/kb/merge', methods=['POST'])
def kb_merge():
    try:
        data = request.json
        side = data.get('side')
        text = (data.get('text') or '').strip()
        if side not in ('firm', 'personal') or not text:
            return jsonify({'success': False, 'error': 'Missing side or text'}), 400
        updated = merge_into_kb_report(side, text, source_label='typed directly into the Knowledge Base page')
        return jsonify({'success': True, 'notes': updated})
    except Exception as e:
        print(f"KB merge error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/kb/notes', methods=['POST'])
def kb_notes_save():
    try:
        data = request.json
        side = data.get('side')
        notes_text = data.get('notes', '')
        if side not in ('firm', 'personal'):
            return jsonify({'success': False, 'error': 'Invalid side'}), 400

        organized = organize_kb_text(side, notes_text)
        notes = load_kb_notes()
        notes[side] = organized
        save_kb_notes(notes)
        return jsonify({'success': True, 'notes': organized})
    except Exception as e:
        print(f"KB notes save error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Single-shot Q&A over the report - no conversation history, just "does the
# report answer this?". Answers strictly from the report text so it can
# honestly say when the knowledge base doesn't have something instead of
# guessing or pulling in outside knowledge.
@app.route('/kb/ask', methods=['POST'])
def kb_ask():
    try:
        data = request.json
        side = data.get('side')
        question = (data.get('question') or '').strip()
        if side not in ('firm', 'personal') or not question:
            return jsonify({'success': False, 'error': 'Missing side or question'}), 400

        report = load_kb_notes().get(side, '').strip()
        connections_context = get_kb_connections_context(side, full=True)
        if not report and not connections_context:
            return jsonify({'success': True, 'answer': "The knowledge base doesn't have any information yet."})

        subject = kb_subject_label(side)
        prompt = (
            f"Here is a knowledge-base report about {subject}:\n\n---\n{report or '(no stored report yet)'}\n---"
            f"{connections_context}\n\n"
            f"Question: {question}\n\n"
            "Answer using ONLY information above (the stored report and/or the live connected files/folders "
            "section, if present). If neither contains the answer, reply with EXACTLY this sentence and nothing "
            "else: \"The knowledge base doesn't have that information.\" Otherwise give a short, direct answer "
            "(one or two sentences, plain text, no markdown) - don't restate the question or pad the answer."
        )
        response = claude_create(
            model=CLAUDE_MODEL,
            max_tokens=300,
            messages=[{'role': 'user', 'content': prompt}]
        )
        answer = "".join(
            block.text for block in response.content if getattr(block, 'type', None) == 'text'
        ).strip()
        return jsonify({'success': True, 'answer': answer})
    except Exception as e:
        print(f"KB ask error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Reads an uploaded file the same way a chat attachment is read (images/PDFs
# go to Claude natively, Office docs are text-extracted first), then asks
# Claude to pull out just the factual bullet points so the Knowledge Base
# gains real facts instead of a wall of raw document text.
@app.route('/kb/upload', methods=['POST'])
def kb_upload():
    try:
        data = request.json
        side = data.get('side')
        filename = data.get('filename', 'file')
        mime_type = data.get('mimeType', '')
        file_data_b64 = data.get('data', '')

        if side not in ('firm', 'personal') or not file_data_b64:
            return jsonify({'success': False, 'error': 'Missing side or file data'}), 400

        blocks = build_attachment_content_blocks([{'name': filename, 'mimeType': mime_type, 'data': file_data_b64}])
        if not blocks:
            return jsonify({'success': False, 'error': "Couldn't read that file"}), 400

        subject = kb_subject_label(side)
        current_report = load_kb_notes().get(side, '').strip()
        instruction = (
            f"You maintain a living knowledge-base report about {subject}. "
            + (f"Here is the CURRENT report:\n\n---\n{current_report}\n---\n\n" if current_report else "There's no report yet. ")
            + f"Extract any relevant facts about {subject} from the attached file (source: \"{filename}\") and reply "
            "with the FULL UPDATED report incorporating them. Keep every still-accurate existing fact, replace "
            "anything the file corrects, and organize clearly with markdown headers (##), bullet points, and a "
            "table where tabular data genuinely fits. Plain factual third-person style, no meta-commentary. "
            "If the file has nothing relevant, reply with the report completely unchanged "
            f"{'(reply with the current report as-is)' if current_report else '(reply with exactly: (no relevant facts found in this file))'}. "
            "Reply with ONLY the report - no preamble, no explanation."
        )
        content_blocks = list(blocks) + [{'type': 'text', 'text': instruction}]

        response = claude_create(
            model=CLAUDE_MODEL,
            max_tokens=1800,
            messages=[{'role': 'user', 'content': content_blocks}]
        )
        updated = "".join(
            block.text for block in response.content if getattr(block, 'type', None) == 'text'
        ).strip()

        notes = load_kb_notes()
        notes[side] = updated
        save_kb_notes(notes)
        return jsonify({'success': True, 'notes': updated})
    except Exception as e:
        print(f"KB upload error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Called right before a Knowledge Base file is permanently deleted - re-reads
# that file's own content and asks Claude to strip out whatever facts came
# from it specifically, so permanent delete actually means the knowledge base
# forgets it (not just that the raw file stops being listed). Archiving a
# file never calls this - only a genuine permanent delete does.
@app.route('/kb/forget', methods=['POST'])
def kb_forget():
    try:
        data = request.json
        side = data.get('side')
        filename = data.get('filename', 'file')
        mime_type = data.get('mimeType', '')
        file_data_b64 = data.get('data', '')

        if side not in ('firm', 'personal') or not file_data_b64:
            return jsonify({'success': False, 'error': 'Missing side or file data'}), 400

        current_report = load_kb_notes().get(side, '').strip()
        if not current_report:
            return jsonify({'success': True, 'notes': ''})

        blocks = build_attachment_content_blocks([{'name': filename, 'mimeType': mime_type, 'data': file_data_b64}])
        if not blocks:
            # Can't re-read the file to know what to remove - leave the report as-is
            # rather than risk deleting unrelated facts.
            return jsonify({'success': True, 'notes': current_report})

        subject = kb_subject_label(side)
        instruction = (
            f"You maintain a living knowledge-base report about {subject}. Here is the CURRENT report:\n\n"
            f"---\n{current_report}\n---\n\n"
            f"The source file attached here (\"{filename}\") is being permanently deleted, and its contribution to "
            "this report must be forgotten. Reply with the FULL UPDATED report with any facts that came ONLY from "
            "this file removed. Keep every fact that is still supported by other information or general knowledge "
            "of the subject. Keep the same markdown structure (headers, bullets, tables) for whatever remains. "
            "If removing this file's facts leaves a section empty, drop that section entirely. "
            "Reply with ONLY the updated report - no preamble, no explanation. If nothing in the report came from "
            "this file, reply with the report completely unchanged."
        )
        content_blocks = list(blocks) + [{'type': 'text', 'text': instruction}]

        response = claude_create(
            model=CLAUDE_MODEL,
            max_tokens=1800,
            messages=[{'role': 'user', 'content': content_blocks}]
        )
        updated = "".join(
            block.text for block in response.content if getattr(block, 'type', None) == 'text'
        ).strip()

        notes = load_kb_notes()
        notes[side] = updated
        save_kb_notes(notes)
        return jsonify({'success': True, 'notes': updated})
    except Exception as e:
        print(f"KB forget error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# --- Knowledge Base: connected files/folders --------------------------
# A connection is a real path on disk (a single file, or a folder - every
# file under it) that the Knowledge Base keeps a live link to. Unlike an
# upload (a one-time snapshot written into the permanent report), a
# connection is never stored in Offload's memory: it's re-checked every time
# the frontend polls /kb/data (and on every chat message, via
# get_kb_connections_context) and its content is summarized fresh, but that
# summary only ever lives on the connection object itself, alongside its
# signature. Disconnect it and the summary is deleted with it - nothing about
# it persists anywhere, and it never touches kb_notes.json.
KB_MAX_FILES_PER_CONNECTION = 40
KB_MAX_FILE_BYTES = 8 * 1024 * 1024
KB_CONNECTION_RAW_TEXT_LIMIT = 600000


def load_kb_connections():
    if os.path.exists(KB_CONNECTIONS_FILE):
        try:
            with open(KB_CONNECTIONS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            data.setdefault('firm', [])
            data.setdefault('personal', [])
            return data
        except (json.JSONDecodeError, OSError):
            pass
    return {'firm': [], 'personal': []}


def save_kb_connections(data):
    with open(KB_CONNECTIONS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


# A lightweight "did anything change" fingerprint - every tracked file's own
# mtime, keyed by its full path. Comparing this dict to the last one caught
# is enough to notice an edit, an added file, or a removed one, without
# having to re-read file contents just to check.
def compute_connection_signature(path, is_folder):
    if not is_folder:
        if not os.path.isfile(path):
            return None
        try:
            return {path: os.path.getmtime(path)}
        except OSError:
            return None

    if not os.path.isdir(path):
        return None
    signature = {}
    for root, dirs, files in os.walk(path):
        dirs[:] = sorted(d for d in dirs if not d.startswith('.'))
        for fname in sorted(files):
            if fname.startswith('.'):
                continue
            fpath = os.path.join(root, fname)
            try:
                if os.path.getsize(fpath) > KB_MAX_FILE_BYTES:
                    continue
                signature[fpath] = os.path.getmtime(fpath)
            except OSError:
                continue
            if len(signature) >= KB_MAX_FILES_PER_CONNECTION:
                return signature
    return signature


# Reads every file named in a signature straight off disk into the same
# {name, mimeType, data} shape a chat attachment arrives in, so the existing
# build_attachment_content_blocks extraction (docx/xlsx/pptx/pdf/images/plain
# text) can be reused as-is instead of duplicating it.
#
# A file open in Excel/Word/etc. often can't be opened directly (Windows
# denies it), even though Explorer-style copies are still permitted - so a
# denied direct read falls back to a shared copy in a temp file before
# giving up on that file. This matters specifically because a "connected"
# file is exactly the kind of file someone keeps open while working in it.
def read_connection_attachments(path, is_folder, signature):
    attachments = []
    for fpath in signature.keys():
        try:
            with open(fpath, 'rb') as f:
                file_bytes = f.read()
        except OSError:
            tmp_path = None
            try:
                fd, tmp_path = tempfile.mkstemp(suffix=os.path.splitext(fpath)[1])
                os.close(fd)
                shutil.copy2(fpath, tmp_path)
                with open(tmp_path, 'rb') as f:
                    file_bytes = f.read()
            except OSError:
                continue
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
        display_name = os.path.relpath(fpath, path) if is_folder else os.path.basename(fpath)
        mime_type = mimetypes.guess_type(fpath)[0] or 'application/octet-stream'
        attachments.append({
            'name': display_name,
            'mimeType': mime_type,
            'data': base64.b64encode(file_bytes).decode('ascii')
        })
    return attachments


# Re-checks one connection against its last-known signature and, only if
# something actually changed, re-reads its current content and regenerates a
# short summary of it. That summary is cached ONLY on the connection object in
# kb_connections.json - it never touches kb_notes.json (the permanent report).
# That's intentional: a connection is a live window onto a real file/folder,
# not something Offload remembers. The moment it's disconnected, its entry
# (summary included) is deleted outright, and nothing about it lingers
# anywhere. Mutates conn in place (signature/missing/lastSyncedAt/summary/
# raw_text) regardless of outcome.
#
# Two different representations get cached, for two different jobs:
#   - raw_text: the file's actual extracted text, verbatim, no LLM involved -
#     this is what chat and /kb/ask actually read from, so a specific figure
#     buried in a big sheet (e.g. "August revenue") is never lost to an LLM's
#     summarization judgment call.
#   - summary: a short LLM-written blurb, used ONLY for the friendly report
#     display (renderKbLiveConnectionsReport) - never for answering questions.
def sync_kb_connection(side, conn):
    if conn.get('type') == 'url':
        return sync_url_connection(side, conn)

    path = conn['path']
    is_folder = conn['type'] == 'folder'

    new_signature = compute_connection_signature(path, is_folder)
    if new_signature is None:
        conn['missing'] = True
        return False
    conn['missing'] = False

    if new_signature == conn.get('signature') and 'summary' in conn and 'raw_text' in conn:
        return False

    attachments = read_connection_attachments(path, is_folder, new_signature)
    if not attachments and new_signature:
        # new_signature is non-empty, so there ARE files that should have
        # been readable - this is a transient failure (e.g. the file is
        # currently open/locked in Excel), not "the content is now empty".
        # Leave signature/summary/raw_text untouched so the last-known-good
        # content keeps being used, and retry on the next sync instead of
        # caching a false "nothing here" result over real data.
        return False

    source_desc = f"connected folder \"{path}\"" if is_folder else f"connected file \"{path}\""
    return _finish_connection_sync(side, conn, attachments, new_signature, source_desc)


# The part both kinds of connection (a path on disk, a link) share once their
# files are in hand: extract the text, keep it for the library, and write the
# short overview.
def _finish_connection_sync(side, conn, attachments, new_signature, source_desc):
    conn['signature'] = new_signature
    conn['lastSyncedAt'] = now_local().isoformat()
    if not attachments:
        conn['summary'] = ''
        conn['raw_text'] = ''
        return True

    blocks = build_attachment_content_blocks(attachments)
    if not blocks:
        conn['summary'] = ''
        conn['raw_text'] = ''
        return True

    raw_text = "\n\n".join(b['text'] for b in blocks if b.get('type') == 'text').strip()
    if len(raw_text) > KB_CONNECTION_RAW_TEXT_LIMIT:
        raw_text = raw_text[:KB_CONNECTION_RAW_TEXT_LIMIT] + "\n\n...(truncated - too large to store in full)"
    conn['raw_text'] = raw_text

    subject = kb_subject_label(side)
    instruction = (
        f"Write a BRIEF, high-level summary of the attached content about {subject} (source: {source_desc}). "
        "This is just an at-a-glance overview shown in a UI panel - the full content is separately available "
        "in full for answering specific questions, so don't try to capture every detail or figure here. "
        "3-5 short markdown bullet points max, covering only the most important, headline-level facts "
        "(e.g. what this source is, the overall totals/bottom line, anything especially notable). No tables, "
        "no sub-sections, no meta-commentary, no preamble. "
        "If there's nothing meaningfully relevant to summarize, reply with exactly: "
        "(no relevant facts found in this connection)"
    )
    # Only the first part of a very large file is needed for an overview.
    overview_blocks = []
    for b in blocks:
        if b.get('type') == 'text' and len(b['text']) > 60000:
            b = {'type': 'text', 'text': b['text'][:60000] + '\n...(rest of file omitted from this overview)'}
        overview_blocks.append(b)
    content_blocks = overview_blocks + [{'type': 'text', 'text': instruction}]

    response = claude_create(
        model=CLAUDE_MODEL,
        max_tokens=400,
        messages=[{'role': 'user', 'content': content_blocks}]
    )
    summary = "".join(
        block.text for block in response.content if getattr(block, 'type', None) == 'text'
    ).strip()
    if summary == '(no relevant facts found in this connection)':
        summary = ''
    conn['summary'] = summary
    return True


# --- Link connections (OneDrive / SharePoint / Google Sheets / any file URL) ---
# A path on this computer can't be read by the hosted web app, so a connection
# can also be a share link: the server downloads the file itself, re-checking at
# most every KB_URL_RECHECK_SECONDS and only re-reading it when the content
# actually changed (daily-updated files just work). Links are fetched with
# guards - https only, and never to a private/internal address.
KB_URL_RECHECK_SECONDS = 300
KB_URL_TIMEOUT_SECONDS = 25


def _is_public_host(hostname):
    try:
        infos = socket.getaddrinfo(hostname, None)
    except OSError:
        return False
    for info in infos:
        ip = ipaddress.ip_address(info[4][0])
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
            return False
    return bool(infos)


class _PublicOnlyRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed = urllib.parse.urlparse(newurl)
        if parsed.scheme != 'https' or not _is_public_host(parsed.hostname or ''):
            raise urllib.error.URLError('Redirected somewhere that is not allowed.')
        return super().redirect_request(req, fp, code, msg, headers, newurl)


# Turns a "share" link into a direct-download one where the service has a
# known pattern; anything else is fetched as given.
def resolve_download_url(url):
    parsed = urllib.parse.urlparse(url)
    host = (parsed.hostname or '').lower()
    if host == '1drv.ms' or host.endswith('.1drv.ms') or host == 'onedrive.live.com':
        encoded = base64.urlsafe_b64encode(url.encode('utf-8')).decode('ascii').rstrip('=')
        return f'https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content'
    if host.endswith('.sharepoint.com'):
        query = dict(urllib.parse.parse_qsl(parsed.query))
        query['download'] = '1'
        return urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(query)))
    sheets = re.match(r'https://docs\.google\.com/spreadsheets/d/([\w-]+)', url)
    if sheets:
        return f'https://docs.google.com/spreadsheets/d/{sheets.group(1)}/export?format=xlsx'
    drive = re.match(r'https://drive\.google\.com/file/d/([\w-]+)', url)
    if drive:
        return f'https://drive.google.com/uc?export=download&id={drive.group(1)}'
    return url


def fetch_url_file(url):
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme != 'https' or not parsed.hostname:
        raise ValueError('The link has to start with https://')
    target = resolve_download_url(url)
    target_host = urllib.parse.urlparse(target).hostname or ''
    if not _is_public_host(target_host):
        raise ValueError("That address can't be reached from here.")
    opener = urllib.request.build_opener(_PublicOnlyRedirectHandler)
    request_obj = urllib.request.Request(target, headers={'User-Agent': 'Mozilla/5.0 (Offload knowledge base)'})
    with opener.open(request_obj, timeout=KB_URL_TIMEOUT_SECONDS) as response:
        data = response.read(KB_MAX_FILE_BYTES + 1)
        if len(data) > KB_MAX_FILE_BYTES:
            raise ValueError(f'That file is larger than {KB_MAX_FILE_BYTES // (1024 * 1024)} MB.')
        content_type = (response.headers.get('Content-Type') or '').split(';')[0].strip().lower()
        disposition = response.headers.get('Content-Disposition') or ''
    filename = ''
    match = re.search(r"filename\*=UTF-8''([^;]+)", disposition) or re.search(r'filename="?([^";]+)"?', disposition)
    if match:
        filename = urllib.parse.unquote(match.group(1))
    if not filename:
        filename = os.path.basename(urllib.parse.unquote(parsed.path)) or 'linked-file'
    if '.' not in os.path.basename(filename):
        filename += mimetypes.guess_extension(content_type) or ''
    if data[:15].lower().startswith((b'<!doctype html', b'<html')):
        raise ValueError("That link opened a web page instead of the file - check it's shared so anyone with the link can view it.")
    return data, filename


def sync_url_connection(side, conn):
    now = now_local()
    last_checked = conn.get('lastCheckedAt')
    if last_checked and 'summary' in conn and 'raw_text' in conn and not conn.get('missing'):
        try:
            if (now - datetime.fromisoformat(last_checked)).total_seconds() < KB_URL_RECHECK_SECONDS:
                return False
        except ValueError:
            pass
    conn['lastCheckedAt'] = now.isoformat()
    try:
        data, filename = fetch_url_file(conn['url'])
    except Exception as e:
        # A failed re-check keeps whatever was last read; only a connection
        # that never worked is flagged as unreachable.
        if 'raw_text' not in conn:
            conn['missing'] = True
        conn['error'] = str(e)
        return False
    conn['missing'] = False
    conn.pop('error', None)

    new_signature = {'sha256': hashlib.sha256(data).hexdigest(), 'size': len(data)}
    if new_signature == conn.get('signature') and 'summary' in conn and 'raw_text' in conn:
        return False

    attachments = [{
        'name': filename,
        'mimeType': mimetypes.guess_type(filename)[0] or 'application/octet-stream',
        'data': base64.b64encode(data).decode('ascii')
    }]
    conn['fileName'] = filename
    return _finish_connection_sync(side, conn, attachments, new_signature, f"linked file \"{filename}\" ({conn['url']})")


# Runs every /kb/data poll (the frontend already refreshes that every ~15s
# while the Knowledge Base page is open) - each connection is a cheap mtime
# check unless something actually changed, in which case that one connection
# gets re-read and merged.
def sync_all_kb_connections():
    connections = load_kb_connections()
    changed = False
    for side in ('firm', 'personal'):
        for conn in connections.get(side, []):
            if sync_kb_connection(side, conn):
                changed = True
    save_kb_connections(connections)
    return connections, changed


@app.route('/kb/connections/add', methods=['POST'])
def kb_connections_add():
    try:
        data = request.json
        side = data.get('side')
        raw_path = (data.get('path') or '').strip().strip('"')
        if side not in ('firm', 'personal') or not raw_path:
            return jsonify({'success': False, 'error': 'Missing side or path'}), 400

        is_link = raw_path.lower().startswith(('http://', 'https://'))
        path = raw_path if is_link else os.path.normpath(raw_path)
        if not is_link and not os.path.exists(path):
            return jsonify({'success': False, 'error': f'No file or folder found at "{path}"'}), 400

        connections = load_kb_connections()
        if any(os.path.normcase(c['path']) == os.path.normcase(path) for c in connections[side]):
            return jsonify({'success': False, 'error': 'That path is already connected'}), 400

        conn = {
            'id': uuid.uuid4().hex,
            'userId': current_user_id(),
            'path': path,
            'type': 'url' if is_link else ('folder' if os.path.isdir(path) else 'file'),
            'addedAt': now_local().isoformat(),
            'signature': {},
            'missing': False
        }
        if is_link:
            conn['url'] = path
            # Prove the link works before keeping it, so a bad or private
            # link is reported now instead of sitting there silently broken.
            try:
                fetch_url_file(path)
            except Exception as link_error:
                return jsonify({'success': False, 'error': f"Couldn't read that link: {link_error}"}), 400
        connections[side].append(conn)
        save_kb_connections(connections)

        # Sync immediately so Francis sees it reflected in the report right
        # away instead of waiting for the next poll.
        sync_kb_connection(side, conn)
        save_kb_connections(connections)

        notes = load_kb_notes()
        return jsonify({'success': True, 'connections': connections, 'notes': notes.get(side, '')})
    except Exception as e:
        print(f"KB connection add error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Stops tracking a connection. Nothing from a connection is ever written into
# kb_notes.json (see sync_kb_connection), so there's no report to "forget" -
# deleting the connection object (summary included) is the whole operation.
@app.route('/kb/connections/remove', methods=['POST'])
def kb_connections_remove():
    try:
        data = request.json
        side = data.get('side')
        conn_id = data.get('id')
        if side not in ('firm', 'personal') or not conn_id:
            return jsonify({'success': False, 'error': 'Missing side or id'}), 400

        connections = load_kb_connections()
        conn = next((c for c in connections[side] if c['id'] == conn_id), None)
        if not conn:
            return jsonify({'success': False, 'error': 'Connection not found'}), 404

        connections[side] = [c for c in connections[side] if c['id'] != conn_id]
        save_kb_connections(connections)

        notes = load_kb_notes().get(side, '')
        return jsonify({'success': True, 'connections': connections, 'notes': notes})
    except Exception as e:
        print(f"KB connection remove error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Opens Windows Explorer at a connected path - a folder opens directly; a
# file opens the folder it's in. Only ever acts on a path that's currently a
# tracked connection, never an arbitrary one the frontend happens to send.
@app.route('/kb/reveal', methods=['POST'])
def kb_reveal():
    try:
        data = request.json
        path = (data.get('path') or '').strip()
        if not path:
            return jsonify({'success': False, 'error': 'Missing path'}), 400

        connections = load_kb_connections()
        known_paths = {os.path.normcase(c['path']) for side in ('firm', 'personal') for c in connections[side]}
        if os.path.normcase(os.path.normpath(path)) not in known_paths:
            return jsonify({'success': False, 'error': 'Not a connected path'}), 400

        if platform.system() != 'Windows':
            return jsonify({'success': False, 'error': 'Revealing in Explorer is only supported on Windows'}), 400

        target = path if os.path.isdir(path) else os.path.dirname(path)
        subprocess.Popen(['explorer', target])
        return jsonify({'success': True})
    except Exception as e:
        print(f"KB reveal error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Opens a native Windows file/folder picker and returns the chosen path -
# used by the Knowledge Base "browse" buttons so the user doesn't have to
# type or paste an absolute path by hand. Runs on the same machine as the
# browser, so a native dialog here is just a local, reversible UI affordance.
@app.route('/kb/browse', methods=['POST'])
def kb_browse():
    try:
        if platform.system() != 'Windows':
            return jsonify({'success': False, 'error': 'Browsing for a file or folder is only supported on Windows'}), 400

        data = request.json or {}
        kind = data.get('type')
        if kind not in ('file', 'folder'):
            return jsonify({'success': False, 'error': 'Invalid type'}), 400

        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        try:
            if kind == 'folder':
                path = filedialog.askdirectory(title='Select a folder to connect', parent=root)
            else:
                path = filedialog.askopenfilename(title='Select a file to connect', parent=root)
        finally:
            root.destroy()

        return jsonify({'success': True, 'path': path or None})
    except Exception as e:
        print(f"KB browse error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# --- Calendar ------------------------------------------------------------
# An internal calendar Ashanti manages: events created here (by Francis, by
# Ashanti via the manage_calendar tool, or pulled in from an external feed)
# live in calendar_events.json. Sync with an outside calendar is one-way in
# each direction over plain ICS - no OAuth, no app registration:
#   - IN:  Francis pastes his external calendar's private ICS feed URL and
#     flips "import" on; that feed gets polled (piggybacking on the existing
#     /calendar/events poll, same pattern as the Knowledge Base connections)
#     and pulled events land here tagged source="external".
#   - OUT: flipping "export" on exposes Offload's own events (source
#     "internal" only - never re-exporting what was just imported, which
#     would loop) as Offload's own ICS feed at a per-install token URL that
#     an external calendar can subscribe to.
CALENDAR_EVENTS_FILE = _data_path('calendar_events.json')
CALENDAR_SETTINGS_FILE = _data_path('calendar_settings.json')
CALENDAR_CHECKIN_STATUS_FILE = _data_path('calendar_checkin_status.json')
AVAILABILITY_SETTINGS_FILE = _data_path('availability_settings.json')
availability_lock = threading.Lock()

# Lunch is deliberately not another fixed Availability block - it's a real
# calendar event, auto-placed fresh each applicable day (see
# _ensure_lunch_generated), that slides within its flexibility window to
# dodge whatever's already on the calendar that day. Once placed it's an
# ordinary event like any other, so every existing free-slot/scheduling path
# already treats it as busy time with no changes needed there.
LUNCH_SETTINGS_FILE = _data_path('lunch_settings.json')
lunch_lock = threading.Lock()

# A single global color for every personal to-do's calendar card (see the
# To-Dos tab in Settings, index.html) - not a per-to-do choice, so there's
# nothing to store on the to-do itself, just this one setting. The 8 options
# are pastel versions of each agent's own brand color (see
# .agent-card[data-agent="..."] background in index.html) - kept in sync
# with TODO_COLOR_PALETTE there.
TODO_SETTINGS_FILE = _data_path('todo_settings.json')
todo_settings_lock = threading.Lock()
TODO_PERSONAL_COLOR_PALETTE = [
    '#a6afc2', '#ffadd9', '#b2d6b2', '#c9b8db', '#ffd7a6', '#eda6a6', '#a6e9ff', '#fff1a6'
]

# Guards every calendar_events.json/calendar_settings.json read-modify-write -
# the frontend's poll, Ashanti's manage_calendar tool, and the checkin poll
# can all land in the same second (threaded=True), and without this two
# concurrent writes can race and silently lose one one's change.
calendar_lock = threading.Lock()


def load_calendar_events():
    if os.path.exists(CALENDAR_EVENTS_FILE):
        try:
            with open(CALENDAR_EVENTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []
    return []


def save_calendar_events(events):
    with open(CALENDAR_EVENTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(events, f, indent=2)


def load_calendar_settings():
    defaults = {
        'importUrl': '',
        'importEnabled': False,
        'exportEnabled': False,
        'exportToken': uuid.uuid4().hex,
        'lastImportSignature': None,
        'lastImportedAt': None,
        'lastImportError': None,
        'dismissedExternalUids': []
    }
    if os.path.exists(CALENDAR_SETTINGS_FILE):
        try:
            with open(CALENDAR_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            defaults.update(data)
            return defaults
        except (json.JSONDecodeError, OSError):
            pass
    save_calendar_settings(defaults)
    return defaults


def save_calendar_settings(data):
    with open(CALENDAR_SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def load_calendar_checkin_status():
    if os.path.exists(CALENDAR_CHECKIN_STATUS_FILE):
        try:
            with open(CALENDAR_CHECKIN_STATUS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_calendar_checkin_status(data):
    with open(CALENDAR_CHECKIN_STATUS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


TODO_CHECKIN_STATUS_FILE = _data_path('todo_checkin_status.json')


def load_todo_checkin_status():
    if os.path.exists(TODO_CHECKIN_STATUS_FILE):
        try:
            with open(TODO_CHECKIN_STATUS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_todo_checkin_status(data):
    with open(TODO_CHECKIN_STATUS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def _unescape_ics_text(s):
    return (s.replace('\\N', '\n').replace('\\n', '\n')
             .replace('\\,', ',').replace('\\;', ';').replace('\\\\', '\\'))


def _escape_ics_text(s):
    return (s or '').replace('\\', '\\\\').replace(';', '\\;').replace(',', '\\,').replace('\n', '\\n')


# Returns (iso_string, is_all_day). Timezone info (TZID params, trailing Z)
# is deliberately not converted - the wall-clock time in the feed is taken
# as-is. Good enough for a single person's own calendars, which are almost
# always all in their own local time anyway.
def _parse_ics_datetime(value):
    value = value.strip()
    try:
        if len(value) == 8 and value.isdigit():
            d = datetime.strptime(value, '%Y%m%d')
            return d.isoformat(), True
        v = value[:-1] if value.endswith('Z') else value
        d = datetime.strptime(v, '%Y%m%dT%H%M%S')
        return d.isoformat(), False
    except ValueError:
        return None, False


# Minimal RFC 5545 reader: unfolds continuation lines, walks VEVENT blocks,
# and reads just the handful of properties a personal calendar actually
# needs (UID/SUMMARY/DTSTART/DTEND/DESCRIPTION/LOCATION/STATUS). Recurring
# events (RRULE) are NOT expanded - only the single instance literally
# written in the feed comes through, same as reading one VEVENT at a time.
def parse_ics(text):
    raw_lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    lines = []
    for line in raw_lines:
        if line.startswith(' ') or line.startswith('\t'):
            if lines:
                lines[-1] += line[1:]
        else:
            lines.append(line)

    events = []
    current = None
    for line in lines:
        stripped = line.strip()
        if stripped == 'BEGIN:VEVENT':
            current = {}
            continue
        if stripped == 'END:VEVENT':
            if current and current.get('start'):
                events.append(current)
            current = None
            continue
        if current is None or ':' not in line:
            continue
        prop, value = line.split(':', 1)
        prop_name = prop.split(';')[0].upper()
        is_date_only = 'VALUE=DATE' in prop.upper() and 'VALUE=DATE-TIME' not in prop.upper()

        if prop_name == 'UID':
            current['externalUid'] = value.strip()
        elif prop_name == 'SUMMARY':
            current['title'] = _unescape_ics_text(value.strip())
        elif prop_name == 'DESCRIPTION':
            current['description'] = _unescape_ics_text(value.strip())
        elif prop_name == 'LOCATION':
            current['location'] = _unescape_ics_text(value.strip())
        elif prop_name == 'DTSTART':
            iso, all_day = _parse_ics_datetime(value)
            if iso:
                current['start'] = iso
                current['allDay'] = is_date_only or all_day
        elif prop_name == 'DTEND':
            iso, _ = _parse_ics_datetime(value)
            if iso:
                current['end'] = iso
        elif prop_name == 'STATUS' and value.strip().upper() == 'CANCELLED':
            current['_cancelled'] = True

    return [e for e in events if not e.get('_cancelled')]


def _fold_ics_line(line):
    if len(line.encode('utf-8')) <= 75:
        return line
    parts = []
    rest = line
    first = True
    while len(rest.encode('utf-8')) > 75:
        cut = 75 if first else 74
        parts.append(rest[:cut])
        rest = rest[cut:]
        first = False
    parts.append(rest)
    return '\r\n '.join(parts)


# The reverse of parse_ics - only ever fed source="internal" events (see
# module comment above), so an imported event is never echoed back out.
def build_ics(events):
    lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Offload//Calendar//EN', 'CALSCALE:GREGORIAN']
    for e in events:
        try:
            start_dt = datetime.fromisoformat(e['start'])
        except (KeyError, ValueError):
            continue
        lines.append('BEGIN:VEVENT')
        lines.append(f"UID:{e['id']}@offload")
        lines.append(f"DTSTAMP:{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}")
        if e.get('allDay'):
            lines.append(f"DTSTART;VALUE=DATE:{start_dt.strftime('%Y%m%d')}")
            try:
                end_dt = datetime.fromisoformat(e['end']) if e.get('end') else start_dt
            except ValueError:
                end_dt = start_dt
            lines.append(f"DTEND;VALUE=DATE:{end_dt.strftime('%Y%m%d')}")
        else:
            lines.append(f"DTSTART:{start_dt.strftime('%Y%m%dT%H%M%S')}")
            if e.get('end'):
                try:
                    end_dt = datetime.fromisoformat(e['end'])
                    lines.append(f"DTEND:{end_dt.strftime('%Y%m%dT%H%M%S')}")
                except ValueError:
                    pass
        lines.append(_fold_ics_line(f"SUMMARY:{_escape_ics_text(e.get('title', ''))}"))
        if e.get('description'):
            lines.append(_fold_ics_line(f"DESCRIPTION:{_escape_ics_text(e['description'])}"))
        if e.get('location'):
            lines.append(_fold_ics_line(f"LOCATION:{_escape_ics_text(e['location'])}"))
        lines.append('END:VEVENT')
    lines.append('END:VCALENDAR')
    return '\r\n'.join(lines) + '\r\n'


# Piggybacks on whatever already polls /calendar/events (the Calendar page
# every ~15s while open, plus every chat turn with Ashanti) rather than
# running its own background thread - same approach as the Knowledge Base's
# file/folder connections. A content hash skips reprocessing when the feed
# hasn't actually changed since last time.
def sync_calendar_import():
    # The network fetch runs unlocked (it can take a few seconds and must
    # never block an event create/update/delete happening at the same time);
    # only the merge-and-save step below needs calendar_lock.
    settings = load_calendar_settings()
    if not settings.get('importEnabled') or not settings.get('importUrl'):
        return settings
    try:
        req = urllib.request.Request(settings['importUrl'], headers={'User-Agent': 'Offload-Calendar/1.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw = resp.read()
        signature = hashlib.sha256(raw).hexdigest()
        if signature == settings.get('lastImportSignature'):
            return settings
        text = raw.decode('utf-8', errors='replace')
        parsed = parse_ics(text)
    except Exception as e:
        settings['lastImportError'] = str(e)
        with calendar_lock:
            save_calendar_settings(settings)
        return settings

    with calendar_lock:
        return _merge_imported_events(parsed, signature)


def _merge_imported_events(parsed, signature):
    try:
        settings = load_calendar_settings()
        events = load_calendar_events()
        dismissed = set(settings.get('dismissedExternalUids', []))
        by_uid = {e.get('externalUid'): e for e in events if e.get('source') == 'external' and e.get('externalUid')}
        seen_uids = set()
        now_iso = now_local().isoformat()

        for pe in parsed:
            uid = pe.get('externalUid')
            if not uid or uid in dismissed:
                continue
            seen_uids.add(uid)
            if uid in by_uid:
                existing = by_uid[uid]
                existing['title'] = pe.get('title', existing.get('title', '(untitled)'))
                existing['description'] = pe.get('description', existing.get('description', ''))
                existing['location'] = pe.get('location', existing.get('location', ''))
                existing['start'] = pe.get('start', existing.get('start'))
                existing['end'] = pe.get('end', existing.get('end'))
                existing['allDay'] = pe.get('allDay', existing.get('allDay', False))
                existing['updatedAt'] = now_iso
            else:
                events.append({
                    'id': uuid.uuid4().hex,
                    'userId': current_user_id(),
                    'title': pe.get('title') or '(untitled)',
                    'description': pe.get('description', ''),
                    'location': pe.get('location', ''),
                    'start': pe.get('start'),
                    'end': pe.get('end') or pe.get('start'),
                    'allDay': pe.get('allDay', False),
                    'source': 'external',
                    'externalUid': uid,
                    'status': 'confirmed',
                    'createdAt': now_iso,
                    'updatedAt': now_iso
                })

        # Anything previously imported whose UID no longer appears in the feed
        # was deleted on the external side - drop it here too.
        events = [
            e for e in events
            if not (e.get('source') == 'external' and e.get('externalUid') and e['externalUid'] not in seen_uids)
        ]

        save_calendar_events(events)
        settings['lastImportSignature'] = signature
        settings['lastImportedAt'] = now_iso
        settings['lastImportError'] = None
    except Exception as e:
        settings['lastImportError'] = str(e)
    save_calendar_settings(settings)
    return settings


CALENDAR_BUSINESS_START_HOUR = 9
CALENDAR_BUSINESS_END_HOUR = 17


# The weekly Availability template (Settings page) - replaces the two fixed
# hour constants above as the actual source of truth for what counts as
# "business" vs "personal" time (and adds a third option, "blackout", that
# was never possible before: time nothing ever gets scheduled into). Each
# day of the week (0=Sunday..6=Saturday, matching JS Date.getDay() so the
# frontend needs no translation) holds an ordered, gapless, non-overlapping
# partition of the full 24 hours - there's no "unset" state, only these
# three types, so the free-slot math below never has to guess what an
# unpainted stretch of time means.
def _default_availability_day():
    return [
        {'start': '00:00', 'end': f'{CALENDAR_BUSINESS_START_HOUR:02d}:00', 'type': 'personal'},
        {'start': f'{CALENDAR_BUSINESS_START_HOUR:02d}:00', 'end': f'{CALENDAR_BUSINESS_END_HOUR:02d}:00', 'type': 'business'},
        {'start': f'{CALENDAR_BUSINESS_END_HOUR:02d}:00', 'end': '24:00', 'type': 'personal'},
    ]


def load_availability_settings():
    if os.path.exists(AVAILABILITY_SETTINGS_FILE):
        try:
            with open(AVAILABILITY_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get('days'), list) and len(data['days']) == 7:
                return data
        except (json.JSONDecodeError, OSError):
            pass
    defaults = {'days': [_default_availability_day() for _ in range(7)]}
    save_availability_settings(defaults)
    return defaults


def save_availability_settings(data):
    with open(AVAILABILITY_SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


# Python's date.weekday() is Monday=0..Sunday=6 - converted once, here,
# to the Sunday=0..Saturday=6 convention this file's availability array and
# the frontend both use, rather than risking a mismatch anywhere else that
# touches day-of-week.
def _day_of_week_sunday0(day):
    return (day.weekday() + 1) % 7


def _minutes_from_hhmm(hhmm):
    h, m = hhmm.split(':')
    return int(h) * 60 + int(m)


# Structural validation only - a day's blocks must exactly tile 00:00-24:00
# with no gaps or overlaps, each with a real type. This is what keeps
# compute_free_slots_for_type honest: it trusts the stored template is a
# clean partition rather than re-checking that on every scheduling call.
def _validate_availability_blocks(blocks):
    if not isinstance(blocks, list) or not blocks:
        return False
    cursor = 0
    for block in blocks:
        if not isinstance(block, dict):
            return False
        if block.get('type') not in ('blackout', 'business', 'personal'):
            return False
        try:
            start = _minutes_from_hhmm(str(block.get('start')))
            end = _minutes_from_hhmm(str(block.get('end')))
        except (ValueError, AttributeError):
            return False
        if start != cursor or end <= start or end > 24 * 60:
            return False
        cursor = end
    return cursor == 24 * 60


# `days` empty means lunch is off entirely - no separate enabled flag, one
# less piece of state that could disagree with itself. lastGeneratedThrough
# is the same rolling-window cursor pattern as a recurring to-do series (see
# _ensure_all_series_generated) - reset to yesterday whenever the settings
# change (see lunch_settings_update) so every not-yet-passed lunch gets
# recomputed under the new rule.
def _default_lunch_settings():
    return {
        'days': [], 'startTime': '12:00', 'lengthMinutes': 60,
        'flexBeforeMinutes': 30, 'flexAfterMinutes': 30,
        'lastGeneratedThrough': (today_local() - timedelta(days=1)).isoformat()
    }


def load_lunch_settings():
    if os.path.exists(LUNCH_SETTINGS_FILE):
        try:
            with open(LUNCH_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get('days'), list):
                return data
        except (json.JSONDecodeError, OSError):
            pass
    defaults = _default_lunch_settings()
    save_lunch_settings(defaults)
    return defaults


def save_lunch_settings(data):
    with open(LUNCH_SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def _validate_lunch_settings(raw):
    try:
        days = sorted({int(d) for d in (raw.get('days') or []) if 0 <= int(d) <= 6})
    except (TypeError, ValueError):
        return None, 'Invalid days'
    start_time = str(raw.get('start_time') or '').strip()
    try:
        _minutes_from_hhmm(start_time)
    except Exception:
        return None, 'Please choose a lunch start time.'
    try:
        length = int(raw.get('length_minutes'))
        flex_before = int(raw.get('flex_before_minutes'))
        flex_after = int(raw.get('flex_after_minutes'))
    except (TypeError, ValueError):
        return None, 'Please fill in the length and flexibility fields.'
    if not (15 <= length <= 4 * 60):
        return None, 'Lunch length must be between 15 minutes and 4 hours.'
    if not (0 <= flex_before <= 4 * 60) or not (0 <= flex_after <= 4 * 60):
        return None, 'Flexibility must be zero or more, up to 4 hours.'
    return {
        'days': days, 'startTime': start_time, 'lengthMinutes': length,
        'flexBeforeMinutes': flex_before, 'flexAfterMinutes': flex_after
    }, None


def _default_todo_settings():
    return {'personalColor': TODO_PERSONAL_COLOR_PALETTE[0]}


def load_todo_settings():
    if os.path.exists(TODO_SETTINGS_FILE):
        try:
            with open(TODO_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict) and data.get('personalColor') in TODO_PERSONAL_COLOR_PALETTE:
                return data
        except (json.JSONDecodeError, OSError):
            pass
    defaults = _default_todo_settings()
    save_todo_settings(defaults)
    return defaults


def save_todo_settings(data):
    with open(TODO_SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def _round_up_to_quarter_hour(dt):
    discard = timedelta(minutes=dt.minute % 15, seconds=dt.second, microseconds=dt.microsecond)
    dt = dt - discard
    if discard:
        dt = dt + timedelta(minutes=15)
    return dt


def _format_time_no_leading_zero(dt):
    return dt.strftime('%I:%M %p').lstrip('0')


# %-d (no leading zero) isn't portable to Windows either - build the label
# from day.day directly instead of relying on a strftime day-of-month code.
def _format_day_label(day):
    return f"{day.strftime('%a %b')} {day.day}"


# The actual open windows within [window_start_minutes, window_end_minutes)
# on `day`, after blocking out every non-cancelled, non-all-day event already
# on the calendar that day - handed to Ashanti so she picks an actual free
# slot instead of doing this interval math herself (which is exactly where
# double-booking and past-the-end-of-day overflow bugs come from). For today
# specifically, the window also can't start before right now. Gaps under 15
# minutes are dropped as not practically bookable. An event an hour or longer
# gets a 15-minute buffer padded onto its end before it blocks out busy time -
# baked in here rather than left as a prompt instruction, so it's guaranteed
# by the math Ashanti already treats as the source of truth, not something
# she has to remember to apply herself. Minute-precision bounds (rather than
# whole hours) so this can be handed one Availability template block at a
# time (see compute_free_slots_for_type) - those can start/end on any
# 15-minute mark, not just the hour.
def compute_free_slots_in_window(day_events, day, now, window_start_minutes, window_end_minutes):
    business_start = datetime.combine(day, datetime.min.time()) + timedelta(minutes=window_start_minutes)
    business_end = datetime.combine(day, datetime.min.time()) + timedelta(minutes=window_end_minutes)
    window_start = business_start
    if day == now.date():
        window_start = max(business_start, _round_up_to_quarter_hour(now))
    if window_start >= business_end:
        return []

    busy = []
    for e in day_events:
        if e.get('allDay'):
            continue
        try:
            s = datetime.fromisoformat(e['start'])
            en = datetime.fromisoformat(e.get('end') or e['start'])
        except (ValueError, KeyError, TypeError):
            continue
        if en - s >= timedelta(hours=1):
            en = en + timedelta(minutes=15)
        s = max(s, window_start)
        en = min(en, business_end)
        if en > s:
            busy.append((s, en))
    busy.sort()

    merged = []
    for s, en in busy:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], en))
        else:
            merged.append((s, en))

    free = []
    cursor = window_start
    for s, en in merged:
        if s > cursor:
            free.append((cursor, s))
        cursor = max(cursor, en)
    if cursor < business_end:
        free.append((cursor, business_end))

    return [(s, en) for s, en in free if (en - s) >= timedelta(minutes=15)]


# Every block of `block_type` ("business" or "personal") on this day's
# Availability template (see load_availability_settings), each run through
# compute_free_slots_in_window and concatenated - a day can have several
# separate business (or personal) blocks (e.g. business 9-12 and 1-5 around
# a blacked-out lunch hour), not just one contiguous window like before.
# Blackout blocks are never matched by either type, so they're correctly
# never free for anything without needing their own special case here.
def compute_free_slots_for_type(day_events, day, now, block_type):
    dow = _day_of_week_sunday0(day)
    day_blocks = load_availability_settings()['days'][dow]
    free = []
    for block in day_blocks:
        if block['type'] != block_type:
            continue
        free.extend(compute_free_slots_in_window(
            day_events, day, now,
            _minutes_from_hhmm(block['start']), _minutes_from_hhmm(block['end'])
        ))
    free.sort()
    return free


def compute_free_slots(day_events, day, now):
    return compute_free_slots_for_type(day_events, day, now, 'business')


# Used only by manage_calendar's move_to_next_available action (see /chat) -
# Ashanti asking to move something "to the next available opening" used to
# mean trusting the model itself to work out a literal start/end datetime
# from the text calendar summary in her prompt, with nothing to actually
# check it against - which is exactly how a request like that once produced
# an event with its end before its start, and a few minutes long instead of
# its real duration. This does what every other scheduling path in this
# file already does instead: real interval math over the actual calendar,
# business hours only, first slot that fits wins, scanning forward day by
# day from `search_from` (clamped to now if it's already past). The model's
# only job becomes resolving words like "next week" into that starting
# point - straightforward for it - not the slot-finding itself.
def _find_next_available_event_slot(events, exclude_event_id, search_from, duration, max_days=60):
    now = now_local()
    if search_from < now:
        search_from = now
    day = search_from.date()
    for _ in range(max_days):
        day_iso = day.isoformat()
        day_events = [
            e for e in events if e['id'] != exclude_event_id and e.get('status') != 'cancelled'
            and not e.get('allDay') and (e.get('start') or '').startswith(day_iso)
        ]
        for s, en in compute_free_slots(day_events, day, now):
            if day == search_from.date() and s < search_from:
                s = search_from
            if en - s >= duration:
                return (s, s + duration)
        day = day + timedelta(days=1)
    return None


# For a personal to-do (see /todos/bulk-schedule) - only the day's blocks
# marked Personal in the Availability template.
def compute_personal_free_slots(day_events, day, now):
    return compute_free_slots_for_type(day_events, day, now, 'personal')


# Whether today's flexibility window for lunch is already over on the clock.
def _lunch_window_passed(day, now, settings):
    natural_start_min = _minutes_from_hhmm(settings['startTime'])
    window_end = min(24 * 60, natural_start_min + settings['lengthMinutes'] + settings['flexAfterMinutes'])
    return day == now.date() and now >= datetime.combine(day, datetime.min.time()) + timedelta(minutes=window_end)


# Where lunch actually lands on `day`, given whatever's already on the
# calendar that day - reuses compute_free_slots_in_window (the same interval
# math every other scheduling path already trusts) over just the flexibility
# window itself, rather than the Availability template's business/personal
# blocks, since lunch is defined purely by its own window, not by hours type.
# Among every free stretch long enough to fit it, picks whichever placement
# sits closest to the configured natural start time - so it stays right at
# that time whenever nothing's in the way, and only drifts earlier/later by
# exactly as much as it has to.
#
# Lunch NEVER leaves its flexibility window (natural start minus "before",
# through natural end plus "after"): if nothing in the window can hold it,
# this returns None and the caller leaves lunch where it already is - it is
# never pushed to some later time of day. `ignore_clock` is for re-checking
# an existing lunch: the clock moving on through today must not by itself
# change where lunch fits (otherwise lunch would drift as the day goes by even
# though nothing on the calendar changed), so slots before "now" still count.
def _find_lunch_slot(day_events, day, now, settings, ignore_clock=False):
    natural_start_min = _minutes_from_hhmm(settings['startTime'])
    length = settings['lengthMinutes']
    window_start = max(0, natural_start_min - settings['flexBeforeMinutes'])
    window_end = min(24 * 60, natural_start_min + length + settings['flexAfterMinutes'])
    day_start = datetime.combine(day, datetime.min.time())

    # For today specifically, once the window has already fully passed on
    # the clock, lunch just doesn't happen today, the same way a to-do never
    # gets scheduled into the past.
    if not ignore_clock and _lunch_window_passed(day, now, settings):
        return None

    fit_now = day_start - timedelta(days=1) if ignore_clock else now

    best = None
    if window_end - window_start >= length:
        for s, en in compute_free_slots_in_window(day_events, day, fit_now, window_start, window_end):
            s_min = (s - day_start).total_seconds() / 60
            en_min = (en - day_start).total_seconds() / 60
            if en_min - s_min < length:
                continue
            candidate = min(max(natural_start_min, s_min), en_min - length)
            if best is None or abs(candidate - natural_start_min) < abs(best - natural_start_min):
                best = candidate
    return day_start + timedelta(minutes=best) if best is not None else None


# Tops up lunch to the same rolling 4-week horizon a recurring to-do series
# uses (see _ensure_all_series_generated) - called from every route that
# either schedules something (todos_bulk_schedule) or just lists todos/
# events, so lunch is always already sitting on the calendar, as a real
# event, before anything else gets a chance to compete for its slot. Once
# generated for a given date, that date's cursor never gets revisited - if
# Francis deletes or moves one day's Lunch card by hand afterward, it stays
# that way rather than quietly reappearing (same reasoning as a recurring
# to-do occurrence never un-deleting itself).
def _ensure_lunch_generated():
    settings = load_lunch_settings()
    if not settings['days']:
        return
    horizon = today_local() + timedelta(days=28)
    if date.fromisoformat(settings['lastGeneratedThrough']) >= horizon:
        return

    with calendar_lock, lunch_lock:
        settings = load_lunch_settings()
        if not settings['days']:
            return
        after = date.fromisoformat(settings['lastGeneratedThrough'])
        if after >= horizon:
            return

        now = now_local()
        working_events = load_calendar_events()
        changed = False
        d = after + timedelta(days=1)
        while d <= horizon:
            if _day_of_week_sunday0(d) in settings['days']:
                day_iso = d.isoformat()
                day_events = [e for e in working_events if e.get('status') != 'cancelled' and (e.get('start') or '').startswith(day_iso)]
                slot_start = _find_lunch_slot(day_events, d, now, settings)
                if slot_start is None and not _lunch_window_passed(d, now, settings):
                    # Nothing in the window is free - lunch still stays inside it, at its natural time.
                    slot_start = datetime.combine(d, datetime.min.time()) + timedelta(minutes=_minutes_from_hhmm(settings['startTime']))
                if slot_start:
                    now_iso = now.isoformat()
                    working_events.append({
                        'id': uuid.uuid4().hex, 'userId': current_user_id(), 'title': 'Lunch',
                        'description': '', 'location': '',
                        'start': slot_start.isoformat(),
                        'end': (slot_start + timedelta(minutes=settings['lengthMinutes'])).isoformat(),
                        'allDay': False, 'source': 'internal', 'origin': 'lunch',
                        'externalUid': None, 'status': 'confirmed',
                        'createdAt': now_iso, 'updatedAt': now_iso
                    })
                    changed = True
            d += timedelta(days=1)

        settings['lastGeneratedThrough'] = horizon.isoformat()
        save_lunch_settings(settings)
        if changed:
            save_calendar_events(working_events)


# Re-derives where Lunch should sit on `day` given whatever's still on the
# calendar, and moves it there (mutating the matching entry in
# `working_events` in place) if that's different from where it currently
# is. Called whenever a real event on that day is deleted (see
# calendar_events_delete) - a lunch shift (see try_shift_lunch_for_slot in
# todos_bulk_schedule) only ever existed to make room for something that
# might now be gone, so it has to be free to revert on its own rather than
# lingering forever once whatever needed it is deleted. Returns True if it
# actually moved anything, so the caller knows whether a save is warranted.
def _reevaluate_lunch_for_day(day, working_events):
    day_iso = day.isoformat()
    lunch_event = next(
        (e for e in working_events if e.get('origin') == 'lunch'
         and e.get('status') != 'cancelled' and (e.get('start') or '').startswith(day_iso)),
        None
    )
    if not lunch_event:
        return False
    lunch_settings = load_lunch_settings()
    if not lunch_settings['days']:
        return False
    now = now_local()
    try:
        current_start = datetime.fromisoformat(lunch_event['start'])
    except (ValueError, KeyError, TypeError):
        return False
    # Lunch that's already underway (or over) today is left alone.
    if day == now.date() and current_start <= now:
        return False
    day_events = [
        e for e in working_events if e is not lunch_event
        and e.get('status') != 'cancelled' and (e.get('start') or '').startswith(day_iso)
    ]
    best_start = _find_lunch_slot(day_events, day, now, lunch_settings, ignore_clock=True)
    if not best_start or best_start.isoformat() == lunch_event['start']:
        return False
    if day == now.date() and best_start <= now:
        return False   # never move lunch back into the past
    length = timedelta(minutes=lunch_settings['lengthMinutes'])
    lunch_event['start'] = best_start.isoformat()
    lunch_event['end'] = (best_start + length).isoformat()
    lunch_event['updatedAt'] = now.isoformat()
    return True


# A broader safety net around _reevaluate_lunch_for_day's DELETE-route hook -
# re-checks every upcoming lunch (not just the one on whatever day a delete
# happened to touch) on every /calendar/events GET, so a shift reverts once
# its cause is gone no matter how that happened - the bulk-schedule popup's
# own removals (red x, blue x, Unplace, drag-off) never touch the backend
# at all until Save, so there's no single delete for that route hook to
# catch; this is what actually snaps lunch back for those, the next time
# anything polls or reloads the calendar. Cheap enough for a single-user
# calendar to just always run rather than trying to track precisely what
# changed.
def _reevaluate_upcoming_lunches(days_ahead=14):
    today = today_local()
    with calendar_lock:
        working_events = load_calendar_events()
        changed = False
        for offset in range(days_ahead):
            if _reevaluate_lunch_for_day(today + timedelta(days=offset), working_events):
                changed = True
        if changed:
            save_calendar_events(working_events)


# Moves that day's Lunch card out of the way of `[busy_start, busy_end)` if
# it's currently overlapping it and can - tries either extreme of its own
# flexibility window (mirrors try_shift_lunch_for_slot in todos_bulk_schedule,
# just driven by an already-known busy range instead of searching for one)
# and only accepts a candidate that clears BOTH the given range and
# everything else real that day. Used when something gets manually dropped
# onto lunch - either a manual drag in the bulk-schedule popup, or a plain
# (non-split-eligible, under 4h) drag on the real Calendar page - since
# neither of those goes through the automatic scheduler's own search.
# Mutates the matching entry in `working_events` in place; returns that
# event on success, or None if there was nothing to do or nothing safe to do.
def _shift_lunch_to_avoid(day, busy_start, busy_end, working_events):
    day_iso = day.isoformat()
    lunch_event = next(
        (e for e in working_events if e.get('origin') == 'lunch'
         and e.get('status') != 'cancelled' and (e.get('start') or '').startswith(day_iso)),
        None
    )
    if not lunch_event:
        return None
    try:
        lunch_s = datetime.fromisoformat(lunch_event['start'])
        lunch_e = datetime.fromisoformat(lunch_event['end'])
    except (ValueError, KeyError, TypeError):
        return None
    if not (lunch_s < busy_end and busy_start < lunch_e):
        return None  # doesn't even overlap - nothing to avoid

    lunch_settings = load_lunch_settings()
    if not lunch_settings['days']:
        return None
    length_min = lunch_settings['lengthMinutes']
    natural_start_min = _minutes_from_hhmm(lunch_settings['startTime'])
    window_start = max(0, natural_start_min - lunch_settings['flexBeforeMinutes'])
    window_end = min(24 * 60, natural_start_min + length_min + lunch_settings['flexAfterMinutes'])
    day_start = datetime.combine(day, datetime.min.time())
    orig_start_min = (lunch_s - day_start).total_seconds() / 60

    other_events = [
        e for e in working_events if e is not lunch_event
        and e.get('status') != 'cancelled' and (e.get('start') or '').startswith(day_iso)
    ]

    def overlaps(a_start, a_end, b_start, b_end):
        return a_start < b_end and b_start < a_end

    def overlaps_others(start_dt, end_dt):
        for e in other_events:
            if e.get('allDay'):
                continue
            try:
                es = datetime.fromisoformat(e['start'])
                ee = datetime.fromisoformat(e.get('end') or e['start'])
            except (ValueError, KeyError, TypeError):
                continue
            if overlaps(es, ee, start_dt, end_dt):
                return True
        return False

    for candidate_min in (window_start, window_end - length_min):
        if candidate_min == orig_start_min:
            continue
        candidate_start = day_start + timedelta(minutes=candidate_min)
        candidate_end = candidate_start + timedelta(minutes=length_min)
        if overlaps(candidate_start, candidate_end, busy_start, busy_end):
            continue
        if overlaps_others(candidate_start, candidate_end):
            continue
        lunch_event['start'] = candidate_start.isoformat()
        lunch_event['end'] = candidate_end.isoformat()
        lunch_event['updatedAt'] = now_local().isoformat()
        return lunch_event

    return None


# Short, chat-context-friendly rundown of what's on the calendar - given to
# Ashanti on every turn so she can talk about it and reference an event's id
# for manage_calendar. Bounded to the next 20 upcoming events so a heavily
# imported calendar doesn't balloon every message.
def get_calendar_context():
    sync_calendar_import()
    events = [e for e in load_calendar_events() if e.get('status') != 'cancelled']
    now = now_local()
    today_iso = now.date().isoformat()
    # The actual clock time, not just the date - without this, "today at
    # 10am" reads as perfectly valid even well after 10am has already
    # passed, since a bare date gives no sense of how far into it we are.
    # strftime's leading-zero-stripping %-I isn't portable to Windows -
    # lstrip('0') gets the same "9:05 AM" instead of "09:05 AM" result.
    now_label = f"today is {today_iso}, current time is {_format_time_no_leading_zero(now)}"
    upcoming = [e for e in events if (e.get('start') or '') >= today_iso]
    upcoming.sort(key=lambda e: e.get('start') or '')
    upcoming = upcoming[:20]

    if upcoming:
        lines = []
        todos_for_ctx = load_todos()
        for e in upcoming:
            start = e.get('start', '')
            when = start[:10] + ' (all day)' if e.get('allDay') else start[:16].replace('T', ' ')
            status_note = f" [{e['status']}]" if e.get('status') not in ('confirmed', None) else ''
            loc_note = f" @ {e['location']}" if e.get('location') else ''
            linked = _todo_for_event(todos_for_ctx, e)
            todo_note = f" [to-do id: {linked['id']}{', completed' if linked.get('completed') else ''}]" if linked else ''
            lines.append(f"- [id: {e['id']}] {when}: {e['title']}{loc_note}{status_note}{todo_note}")
        events_block = f"\n\nCALENDAR ({now_label}):\n" + "\n".join(lines)
    else:
        events_block = f"\n\nCALENDAR ({now_label}): Nothing on Francis's Offload calendar right now."

    # Pre-computed free windows for the next 7 days, already accounting for
    # business hours and every existing event - this is the actual source of
    # truth for "is there room", not something to re-derive from the raw
    # event list above.
    free_lines = []
    for offset in range(7):
        day = now.date() + timedelta(days=offset)
        day_iso = day.isoformat()
        day_events = [e for e in events if (e.get('start') or '').startswith(day_iso)]
        slots = compute_free_slots(day_events, day, now)
        if offset == 0:
            day_label = f"Today ({_format_day_label(day)})"
        elif offset == 1:
            day_label = f"Tomorrow ({_format_day_label(day)})"
        else:
            day_label = _format_day_label(day)
        if slots:
            slots_text = ', '.join(
                f"{_format_time_no_leading_zero(s)}-{_format_time_no_leading_zero(en)}" for s, en in slots
            )
            free_lines.append(f"- {day_label}: {slots_text}")
        else:
            free_lines.append(f"- {day_label}: fully booked, no room left in business hours")
    free_block = (
        "\n\nFREE TIME (within Francis's configured business hours for each day - see the "
        "Availability settings - already accounts for every event above and, "
        "for today, the current time) - when scheduling something, only ever use a window from "
        "here that's at least as long as the event's duration; never pick a time this doesn't "
        "list as free:\n" + "\n".join(free_lines)
    )

    return events_block + free_block


@app.route('/calendar/events', methods=['GET'])
def calendar_events_list():
    _ensure_lunch_generated()
    _reevaluate_upcoming_lunches()
    settings = sync_calendar_import()
    events = load_calendar_events()
    events.sort(key=lambda e: e.get('start') or '')
    public_settings = {k: v for k, v in settings.items() if k != 'dismissedExternalUids'}
    public_settings['exportUrl'] = (
        request.host_url.rstrip('/') + f"/calendar/export/{settings['exportToken']}.ics"
        if settings.get('exportEnabled') else None
    )
    return jsonify({'success': True, 'events': events, 'settings': public_settings})


@app.route('/calendar/events', methods=['POST'])
def calendar_events_create():
    try:
        data = request.json or {}
        title = str(data.get('title', '')).strip()
        start = str(data.get('start', '')).strip()
        if not title or not start:
            return jsonify({'success': False, 'error': 'Missing title or start'}), 400
        now_iso = now_local().isoformat()
        event = {
            'id': uuid.uuid4().hex,
            'userId': current_user_id(),
            'title': title,
            'description': str(data.get('description', '') or '').strip(),
            'location': str(data.get('location', '') or '').strip(),
            'start': start,
            'end': str(data.get('end', '') or '').strip() or start,
            'allDay': bool(data.get('allDay', False)),
            'source': 'internal',
            'origin': 'manual',
            'externalUid': None,
            'status': 'confirmed',
            'createdAt': now_iso,
            'updatedAt': now_iso,
            'attachments': data.get('attachments') or []
        }
        with calendar_lock:
            events = load_calendar_events()
            events.append(event)
            save_calendar_events(events)
        return jsonify({'success': True, 'event': event})
    except Exception as e:
        print(f"Calendar create error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/calendar/events/<event_id>', methods=['POST'])
def calendar_events_update(event_id):
    try:
        data = request.json or {}
        with calendar_lock:
            events = load_calendar_events()
            event = next((e for e in events if e['id'] == event_id), None)
            if not event:
                return jsonify({'success': False, 'error': 'Event not found'}), 404
            for field in ('title', 'description', 'location', 'start', 'end'):
                if field in data:
                    event[field] = str(data[field] or '').strip()
            if 'allDay' in data:
                event['allDay'] = bool(data['allDay'])
            if 'status' in data and data['status'] in ('confirmed', 'done', 'cancelled'):
                event['status'] = data['status']
            if 'attachments' in data:
                event['attachments'] = data['attachments'] or []
            event['updatedAt'] = now_local().isoformat()
            save_calendar_events(events)
        return jsonify({'success': True, 'event': event})
    except Exception as e:
        print(f"Calendar update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Splits an existing event into two, in place - used when dragging a 4h+
# to-do card over lunch or past the end of the day on the real Calendar
# page (see the drag-split logic in index.html: computeCalendarDragSplit,
# onCalendarEventDragEnd). The event being split becomes part 1 (keeps its
# id, so a todo's calendarEventId never has to change); part 2 is a new
# event carrying the same title/description/location/origin, linked back
# via splitOf - the exact same shape try_split_around_lunch produces for
# the automatic scheduler, so every existing split-aware code path
# (findTodoForCalendarEvent, delete/unschedule cascades, the "(1/2)"
# display) already handles it with no changes needed here.
@app.route('/calendar/events/<event_id>/split', methods=['POST'])
def calendar_events_split(event_id):
    try:
        data = request.json or {}
        try:
            part1_start = str(data['part1_start'])
            part1_end = str(data['part1_end'])
            part2_start = str(data['part2_start'])
            part2_end = str(data['part2_end'])
        except KeyError:
            return jsonify({'success': False, 'error': 'Missing split part times'}), 400
        with calendar_lock:
            events = load_calendar_events()
            event = next((e for e in events if e['id'] == event_id), None)
            if not event:
                return jsonify({'success': False, 'error': 'Event not found'}), 404
            now_iso = now_local().isoformat()
            event['start'] = part1_start
            event['end'] = part1_end
            event['splitPart'] = 1
            event['splitTotal'] = 2
            event['updatedAt'] = now_iso
            event2 = dict(event)
            event2['id'] = uuid.uuid4().hex
            event2['start'] = part2_start
            event2['end'] = part2_end
            event2['splitPart'] = 2
            event2['splitOf'] = event['id']
            event2['createdAt'] = now_iso
            event2['updatedAt'] = now_iso
            events.append(event2)
            try:
                _reevaluate_lunch_for_day(datetime.fromisoformat(part1_start).date(), events)
                _reevaluate_lunch_for_day(datetime.fromisoformat(part2_start).date(), events)
            except (ValueError, TypeError):
                pass
            save_calendar_events(events)
        return jsonify({'success': True, 'event': event, 'event2': event2})
    except Exception as e:
        print(f"Calendar split error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/calendar/events/<event_id>', methods=['DELETE'])
def calendar_events_delete(event_id):
    try:
        with calendar_lock:
            events = load_calendar_events()
            event = next((e for e in events if e['id'] == event_id), None)
            if not event:
                return jsonify({'success': False, 'error': 'Event not found'}), 404
            events = [e for e in events if e['id'] != event_id]
            if not event.get('allDay') and event.get('start'):
                try:
                    _reevaluate_lunch_for_day(datetime.fromisoformat(event['start']).date(), events)
                except (ValueError, TypeError):
                    pass
            save_calendar_events(events)
            # An imported event that's deleted locally would just reappear on
            # the next sync otherwise - remember its UID so re-import skips it.
            if event.get('source') == 'external' and event.get('externalUid'):
                settings = load_calendar_settings()
                dismissed = set(settings.get('dismissedExternalUids', []))
                dismissed.add(event['externalUid'])
                settings['dismissedExternalUids'] = list(dismissed)
                save_calendar_settings(settings)
        _delete_item_attachments(event)
        return jsonify({'success': True})
    except Exception as e:
        print(f"Calendar delete error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/calendar/settings', methods=['POST'])
def calendar_settings_update():
    try:
        data = request.json or {}
        with calendar_lock:
            settings = load_calendar_settings()
            if 'importUrl' in data:
                new_url = str(data['importUrl'] or '').strip()
                if new_url != settings.get('importUrl'):
                    settings['lastImportSignature'] = None  # force a fresh pull on URL change
                settings['importUrl'] = new_url
            if 'importEnabled' in data:
                settings['importEnabled'] = bool(data['importEnabled'])
            if 'exportEnabled' in data:
                settings['exportEnabled'] = bool(data['exportEnabled'])
            save_calendar_settings(settings)
        if settings['importEnabled']:
            settings = sync_calendar_import()
        public_settings = {k: v for k, v in settings.items() if k != 'dismissedExternalUids'}
        public_settings['exportUrl'] = (
            request.host_url.rstrip('/') + f"/calendar/export/{settings['exportToken']}.ics"
            if settings.get('exportEnabled') else None
        )
        return jsonify({'success': True, 'settings': public_settings})
    except Exception as e:
        print(f"Calendar settings error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Settings page > Availability - the weekly blackout/business/personal
# template that compute_free_slots/compute_personal_free_slots actually
# schedule against (see those functions above). The frontend always paints
# and sends back one whole day's complete block list at a time (see
# spliceAvailabilityBlock in index.html), so this route's job is just to
# validate that's a clean partition and persist it, not to compute the
# splice itself.
@app.route('/availability', methods=['GET'])
def availability_get():
    return jsonify({'success': True, 'availability': load_availability_settings()})


@app.route('/availability', methods=['POST'])
def availability_update():
    try:
        data = request.json or {}
        day_of_week = data.get('day_of_week')
        blocks = data.get('blocks')
        if not isinstance(day_of_week, int) or not (0 <= day_of_week <= 6):
            return jsonify({'success': False, 'error': 'Invalid day_of_week'}), 400
        if not _validate_availability_blocks(blocks):
            return jsonify({'success': False, 'error': 'Blocks must be a gapless, non-overlapping partition of the day'}), 400
        with availability_lock:
            settings = load_availability_settings()
            settings['days'][day_of_week] = blocks
            save_availability_settings(settings)
        return jsonify({'success': True, 'availability': settings})
    except Exception as e:
        print(f"Availability update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/calendar/lunch-settings', methods=['GET'])
def lunch_settings_get():
    return jsonify({'success': True, 'settings': load_lunch_settings()})


@app.route('/calendar/lunch-settings', methods=['POST'])
def lunch_settings_update():
    try:
        validated, err = _validate_lunch_settings(request.json or {})
        if err:
            return jsonify({'success': False, 'error': err}), 400
        with lunch_lock, calendar_lock:
            settings = load_lunch_settings()
            settings.update(validated)
            # Every not-yet-passed Lunch card (today or later) gets cleared
            # so the next generation pass recreates it under the new rule -
            # past ones are left alone as history, same "future only" scope
            # as a recurring to-do's "apply to all" edit.
            today_iso = today_local().isoformat()
            working_events = [
                e for e in load_calendar_events()
                if not (e.get('origin') == 'lunch' and (e.get('start') or '') >= today_iso)
            ]
            save_calendar_events(working_events)
            settings['lastGeneratedThrough'] = (today_local() - timedelta(days=1)).isoformat()
            save_lunch_settings(settings)
        _ensure_lunch_generated()
        return jsonify({'success': True, 'settings': load_lunch_settings()})
    except Exception as e:
        print(f"Lunch settings update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/todos/settings', methods=['GET'])
def todo_settings_get():
    return jsonify({'success': True, 'settings': load_todo_settings(), 'palette': TODO_PERSONAL_COLOR_PALETTE})


@app.route('/todos/settings', methods=['POST'])
def todo_settings_update():
    try:
        color = str((request.json or {}).get('personal_color') or '').strip()
        if color not in TODO_PERSONAL_COLOR_PALETTE:
            return jsonify({'success': False, 'error': 'Invalid color'}), 400
        with todo_settings_lock:
            settings = load_todo_settings()
            settings['personalColor'] = color
            save_todo_settings(settings)
        return jsonify({'success': True, 'settings': settings})
    except Exception as e:
        print(f"Todo settings update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Called when something gets manually dropped onto lunch - see
# _shift_lunch_to_avoid for the actual move logic. `start`/`end` describe
# whatever needs the room; the response says whether lunch actually moved
# (it may not need to, or may not be able to) and, if so, where.
@app.route('/calendar/lunch-settings/shift', methods=['POST'])
def lunch_settings_shift():
    try:
        data = request.json or {}
        try:
            busy_start = datetime.fromisoformat(str(data['start']))
            busy_end = datetime.fromisoformat(str(data['end']))
        except (KeyError, ValueError, TypeError):
            return jsonify({'success': False, 'error': 'Invalid start/end'}), 400
        with calendar_lock:
            working_events = load_calendar_events()
            moved_event = _shift_lunch_to_avoid(busy_start.date(), busy_start, busy_end, working_events)
            if moved_event:
                save_calendar_events(working_events)
        return jsonify({'success': True, 'moved': bool(moved_event), 'event': moved_event})
    except Exception as e:
        print(f"Lunch shift error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Public-by-obscurity feed URL (the token is the only guard, same security
# model every consumer calendar uses for a "private address") - only ever
# serves source="internal" events, so an imported event is never echoed
# back out to where it came from.
@app.route('/calendar/export/<token>.ics')
def calendar_export(token):
    settings = load_calendar_settings()
    if not settings.get('exportEnabled') or token != settings.get('exportToken'):
        return ('Not found', 404)
    events = [e for e in load_calendar_events() if e.get('source') == 'internal' and e.get('status') != 'cancelled']
    return Response(build_ics(events), mimetype='text/calendar')


# Ashanti "popping in" - polled by the frontend every few minutes while the
# app is open. Cheap on every call (just comparing today's event times to
# now); only spends an LLM call the moment something is actually due, and
# calendar_checkin_status.json remembers what's already been shown today so
# the same nudge doesn't repeat on the next poll.
@app.route('/calendar/checkins/today', methods=['GET'])
def calendar_checkins_today():
    try:
        sync_calendar_import()
        today_iso = today_local().isoformat()
        todays_events = [
            e for e in load_calendar_events()
            if not e.get('allDay') and e.get('status') == 'confirmed' and (e.get('start') or '').startswith(today_iso)
        ]
        todays_events.sort(key=lambda e: e['start'])

        now = now_local()
        status = load_calendar_checkin_status()
        # Trim old days so this file doesn't grow forever.
        status = {k: v for k, v in status.items() if k >= (today_local() - timedelta(days=2)).isoformat()}
        shown_today = status.get(today_iso, [])

        item = None
        for e in todays_events:
            try:
                start_dt = datetime.fromisoformat(e['start'])
            except ValueError:
                continue
            minutes_until = (start_dt - now).total_seconds() / 60
            upcoming_key = f"{e['id']}:upcoming"
            if 0 <= minutes_until <= 15 and upcoming_key not in shown_today:
                item = {'key': upcoming_key, 'kind': 'upcoming', 'event': e, 'minutes_until': round(minutes_until)}
                break
            end_dt = None
            if e.get('end'):
                try:
                    end_dt = datetime.fromisoformat(e['end'])
                except ValueError:
                    end_dt = None
            overdue_key = f"{e['id']}:overdue"
            if end_dt and now > end_dt and overdue_key not in shown_today:
                item = {'key': overdue_key, 'kind': 'overdue', 'event': e}
                break

        if not item:
            save_calendar_checkin_status(status)
            return jsonify({'success': True, 'checkin': None})

        e = item['event']
        if item['kind'] == 'upcoming':
            prompt = (
                f"His event \"{e['title']}\" starts in about {item['minutes_until']} minutes"
                + (f" at {e['location']}" if e.get('location') else "") + ". Write one short, warm, "
                "in-character check-in (1-2 sentences) reminding him and asking if he's still on track or "
                "needs to adjust the plan. No preamble, no quotation marks around it - just the message."
            )
        else:
            prompt = (
                f"His event \"{e['title']}\" was supposed to have ended by now but is still open on the "
                "calendar. Write one short, warm, in-character check-in (1-2 sentences) asking if it's done "
                "or if the rest of the day's plan needs to shift. No preamble, no quotation marks around it "
                "- just the message."
            )

        try:
            response = claude_create(
                model=CLAUDE_MODEL,
                max_tokens=150,
                log_agent='ashanti',
                system=get_agent_system_prompt('ashanti'),
                messages=[{'role': 'user', 'content': prompt}]
            )
            message = "".join(b.text for b in response.content if getattr(b, 'type', None) == 'text').strip()
        except Exception:
            message = f"Quick check-in: \"{e['title']}\" is coming up soon - still on track?"

        shown_today.append(item['key'])
        status[today_iso] = shown_today
        save_calendar_checkin_status(status)

        return jsonify({'success': True, 'checkin': {'event_id': e['id'], 'message': message}})
    except Exception as ex:
        print(f"Calendar checkin error: {ex}")
        return jsonify({'success': False, 'error': str(ex)}), 500


# --- Discussion Topics ---------------------------------------------------
# An ongoing, user-maintained list of things to bring up with the team,
# grouped into categories - each category has one agent responsible for
# that subject matter, so every topic under it shares that same "Discuss
# with X" target. Marking a topic discussed doesn't delete it - it just
# moves it into the grayed-out/struck-through section (still with its
# Discuss button live, in case it needs to come back up).
DISCUSSION_TOPICS_FILE = _data_path('discussion_topics.json')
discussion_topics_lock = threading.Lock()


def load_discussion_topics():
    if os.path.exists(DISCUSSION_TOPICS_FILE):
        try:
            with open(DISCUSSION_TOPICS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            data.setdefault('categories', [])
            return data
        except (json.JSONDecodeError, OSError):
            pass
    return {'categories': []}


def save_discussion_topics(data):
    with open(DISCUSSION_TOPICS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def find_topic_category(data, topic_id):
    for cat in data['categories']:
        for topic in cat.get('topics', []):
            if topic['id'] == topic_id:
                return cat, topic
    return None, None


# The page is organized by agent, so a topic is filed under an agent and its
# category is only a label (shown as a small tag). Finds that agent's category
# with this name (creating it if needed); with no name it uses "General".
def get_or_create_agent_category(store, agent, name=None):
    wanted = (name or 'General').strip() or 'General'
    for cat in store['categories']:
        if cat['agent'] == agent and cat['name'].strip().lower() == wanted.lower():
            return cat
    category = {
        'id': uuid.uuid4().hex,
        'userId': current_user_id(),
        'name': wanted,
        'agent': agent,
        'createdAt': now_local().isoformat(),
        'topics': []
    }
    store['categories'].append(category)
    return category


# Pulls just the "Core Role" blurb out of an agent's system prompt (skipping
# personality traits, formatting rules, etc.) so the categorizer below has
# enough to match a topic's subject matter without a giant prompt.
def get_agent_role_summary(agent):
    try:
        with open(f'agents/{agent}/system_prompt.txt', 'r', encoding='utf-8') as f:
            text = f.read()
    except FileNotFoundError:
        return ''
    marker = '## Core Role'
    idx = text.find(marker)
    if idx == -1:
        return ''
    after = text[idx + len(marker):]
    next_heading = after.find('\n##')
    chunk = after[:next_heading] if next_heading != -1 else after
    return chunk.strip()


def build_agent_roster_text():
    return "\n".join(f"- {agent}: {get_agent_role_summary(agent)}" for agent in ALL_AGENTS)


# Decides where a new topic belongs: an existing category if one genuinely
# fits, otherwise a brand-new category name plus whichever team member's
# actual domain (per their Core Role) best matches the topic. Pure
# classification, no file I/O - callers do the actual read-modify-write
# themselves, under discussion_topics_lock, so this can run without holding
# it (an LLM call is exactly the kind of slow operation that lock must never
# be held across - see the calendar import sync for the same lesson).
def categorize_discussion_topic(text, details, existing_categories):
    existing_list = "\n".join(
        f"- \"{c['name']}\" (handled by {c['agent']})" for c in existing_categories
    ) or "(none yet)"
    roster = build_agent_roster_text()
    topic_desc = f"\"{text}\"" + (f"\n\nDetails:\n{details}" if details else "")
    prompt = (
        f"A new discussion topic needs to be filed into a category: {topic_desc}\n\n"
        f"Existing categories:\n{existing_list}\n\n"
        f"Team roster (for picking who should own a brand-new category):\n{roster}\n\n"
        "If this topic clearly belongs in one of the existing categories, reply with exactly:\n"
        "CATEGORY: <the existing category's exact name>\n"
        "AGENT: existing\n\n"
        "If it doesn't fit any existing category well, invent a short, clear new category name (2-4 words, "
        "title case, describing the subject matter - not the agent's name) and pick the ONE team member "
        "from the roster whose actual domain best matches this topic. Reply with exactly:\n"
        "CATEGORY: <new category name>\n"
        "AGENT: <agent id from the roster>\n\n"
        "No other text, no explanation."
    )
    response = claude_create(
        model=CLAUDE_MODEL,
        max_tokens=150,
        messages=[{'role': 'user', 'content': prompt}]
    )
    reply = "".join(b.text for b in response.content if getattr(b, 'type', None) == 'text').strip()
    category_name, agent = None, None
    for line in reply.splitlines():
        line = line.strip()
        if line.upper().startswith('CATEGORY:'):
            category_name = line.split(':', 1)[1].strip().strip('"')
        elif line.upper().startswith('AGENT:'):
            agent = line.split(':', 1)[1].strip().lower()
    return category_name, agent


@app.route('/discussion-topics/auto-add', methods=['POST'])
def discussion_topics_auto_add():
    try:
        data = request.json or {}
        text = str(data.get('text', '')).strip()
        details = str(data.get('details', '') or '').strip()
        if not text or not details:
            return jsonify({'success': False, 'error': 'Missing text or details'}), 400

        existing_snapshot = load_discussion_topics()['categories']
        category_name, agent = categorize_discussion_topic(text, details, existing_snapshot)

        with discussion_topics_lock:
            store = load_discussion_topics()
            matched = None
            if category_name:
                matched = next(
                    (c for c in store['categories'] if c['name'].strip().lower() == category_name.strip().lower()),
                    None
                )

            is_new_category = matched is None
            if matched:
                category = matched
            else:
                category = {
                    'id': uuid.uuid4().hex,
                    'userId': current_user_id(),
                    'name': category_name or 'General',
                    'agent': agent if agent in ALL_AGENTS else ALL_AGENTS[0],
                    'createdAt': now_local().isoformat(),
                    'topics': []
                }
                store['categories'].append(category)

            topic = {
                'id': uuid.uuid4().hex,
                'userId': current_user_id(),
                'text': text,
                'details': details,
                'discussed': False,
                'createdAt': now_local().isoformat(),
                'discussedAt': None,
                'attachments': data.get('attachments') or []
            }
            category['topics'].append(topic)
            save_discussion_topics(store)
            result = {
                'success': True,
                'data': store,
                'category_id': category['id'],
                'topic_id': topic['id'],
                'category_name': category['name'],
                'is_new_category': is_new_category
            }
        return jsonify(result)
    except Exception as e:
        print(f"Discussion topics auto-add error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/discussion-topics', methods=['GET'])
def discussion_topics_list():
    return jsonify({'success': True, 'data': load_discussion_topics()})


@app.route('/discussion-topics/categories', methods=['POST'])
def discussion_topics_create_category():
    try:
        data = request.json or {}
        name = str(data.get('name', '')).strip()
        agent = data.get('agent')
        if not name or agent not in ALL_AGENTS:
            return jsonify({'success': False, 'error': 'Missing name or invalid agent'}), 400
        with discussion_topics_lock:
            store = load_discussion_topics()
            category = {
                'id': uuid.uuid4().hex,
                'userId': current_user_id(),
                'name': name,
                'agent': agent,
                'createdAt': now_local().isoformat(),
                'topics': []
            }
            store['categories'].append(category)
            save_discussion_topics(store)
        return jsonify({'success': True, 'data': store})
    except Exception as e:
        print(f"Discussion topics create category error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/discussion-topics/categories/<category_id>', methods=['POST'])
def discussion_topics_update_category(category_id):
    try:
        data = request.json or {}
        with discussion_topics_lock:
            store = load_discussion_topics()
            category = next((c for c in store['categories'] if c['id'] == category_id), None)
            if not category:
                return jsonify({'success': False, 'error': 'Category not found'}), 404
            if 'name' in data:
                new_name = str(data['name'] or '').strip()
                if new_name:
                    category['name'] = new_name
            if 'agent' in data and data['agent'] in ALL_AGENTS:
                category['agent'] = data['agent']
            save_discussion_topics(store)
        return jsonify({'success': True, 'data': store})
    except Exception as e:
        print(f"Discussion topics update category error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/discussion-topics/categories/<category_id>', methods=['DELETE'])
def discussion_topics_delete_category(category_id):
    try:
        with discussion_topics_lock:
            store = load_discussion_topics()
            if not any(c['id'] == category_id for c in store['categories']):
                return jsonify({'success': False, 'error': 'Category not found'}), 404
            store['categories'] = [c for c in store['categories'] if c['id'] != category_id]
            save_discussion_topics(store)
        return jsonify({'success': True, 'data': store})
    except Exception as e:
        print(f"Discussion topics delete category error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/discussion-topics/topics', methods=['POST'])
def discussion_topics_create_topic():
    try:
        data = request.json or {}
        category_id = data.get('category_id')
        agent = data.get('agent')
        text = str(data.get('text', '')).strip()
        details = str(data.get('details', '') or '').strip()
        if not text or not (category_id or agent in ALL_AGENTS):
            return jsonify({'success': False, 'error': 'Missing text, and either a category or an agent'}), 400
        with discussion_topics_lock:
            store = load_discussion_topics()
            if category_id:
                category = next((c for c in store['categories'] if c['id'] == category_id), None)
                if not category:
                    return jsonify({'success': False, 'error': 'Category not found'}), 404
            else:
                category = get_or_create_agent_category(store, agent)
            topic = {
                'id': uuid.uuid4().hex,
                'userId': current_user_id(),
                'text': text,
                'details': details,
                'discussed': False,
                'createdAt': now_local().isoformat(),
                'discussedAt': None,
                'attachments': data.get('attachments') or []
            }
            category['topics'].append(topic)
            save_discussion_topics(store)
        return jsonify({'success': True, 'data': store, 'category_id': category['id'], 'topic_id': topic['id']})
    except Exception as e:
        print(f"Discussion topics create topic error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/discussion-topics/topics/<topic_id>', methods=['POST'])
def discussion_topics_update_topic(topic_id):
    try:
        data = request.json or {}
        with discussion_topics_lock:
            store = load_discussion_topics()
            current_category, topic = find_topic_category(store, topic_id)
            if not topic:
                return jsonify({'success': False, 'error': 'Topic not found'}), 404
            if 'text' in data:
                new_text = str(data['text'] or '').strip()
                if new_text:
                    topic['text'] = new_text
            if 'details' in data:
                topic['details'] = str(data['details'] or '').strip()
            if 'discussed' in data:
                topic['discussed'] = bool(data['discussed'])
                topic['discussedAt'] = now_local().isoformat() if topic['discussed'] else None
            if 'attachments' in data:
                topic['attachments'] = data['attachments'] or []
            if data.get('agent') in ALL_AGENTS and data['agent'] != current_category['agent']:
                # Handing it to another agent: it keeps its category label there.
                new_category = get_or_create_agent_category(store, data['agent'], current_category['name'])
                current_category['topics'] = [t for t in current_category['topics'] if t['id'] != topic_id]
                new_category['topics'].append(topic)
                if not current_category['topics']:
                    store['categories'] = [c for c in store['categories'] if c['id'] != current_category['id']]
            elif 'category_id' in data and data['category_id'] and data['category_id'] != current_category['id']:
                new_category = next((c for c in store['categories'] if c['id'] == data['category_id']), None)
                if not new_category:
                    return jsonify({'success': False, 'error': 'Category not found'}), 404
                current_category['topics'] = [t for t in current_category['topics'] if t['id'] != topic_id]
                new_category['topics'].append(topic)
            save_discussion_topics(store)
        return jsonify({'success': True, 'data': store})
    except Exception as e:
        print(f"Discussion topics update topic error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/discussion-topics/topics/<topic_id>', methods=['DELETE'])
def discussion_topics_delete_topic(topic_id):
    try:
        with discussion_topics_lock:
            store = load_discussion_topics()
            category, topic = find_topic_category(store, topic_id)
            if not topic:
                return jsonify({'success': False, 'error': 'Topic not found'}), 404
            category['topics'] = [t for t in category['topics'] if t['id'] != topic_id]
            save_discussion_topics(store)
        _delete_item_attachments(topic)
        return jsonify({'success': True, 'data': store})
    except Exception as e:
        print(f"Discussion topics delete topic error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# --- To-Do List -----------------------------------------------------------
# A flat personal to-do list - no categories or agents involved, unlike
# Discussion Topics. Each item just has a title and its details (one point
# per line, rendered as bullets - same convention as Discussion Topics).
TODOS_FILE = _data_path('todos.json')
todos_lock = threading.Lock()

# A recurring to-do's template + rule (see the /todos/series routes below) -
# each occurrence it spawns is a completely ordinary todo/calendar-event pair
# (with a `seriesId`/`occurrenceDate` added) so every existing todo/event code
# path - complete, reschedule, delete, drag, the Edit Event modal - already
# works on it unchanged. This file only ever holds the series definitions
# themselves, never the occurrences.
TODO_SERIES_FILE = _data_path('todo_series.json')
todo_series_lock = threading.Lock()

# A reusable starting point for a new to-do (see Settings > To-Dos, and the
# "Use a template" picker on the To-Do page) - just the fields a fresh to-do
# itself needs (title/details/estimated time/priority/personal), nothing
# about scheduling. Picking one only ever fills in the Add To-Do form for
# Francis to review/adjust before submitting - it never creates a to-do by
# itself.
TODO_TEMPLATES_FILE = _data_path('todo_templates.json')
todo_templates_lock = threading.Lock()


def load_todos():
    if os.path.exists(TODOS_FILE):
        try:
            with open(TODOS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                return _apply_todo_rollovers(data)
        except (json.JSONDecodeError, OSError):
            pass
    return []


# A to-do can be scheduled more than once: calendarEventIds is every calendar
# event it has ever been scheduled onto (oldest first), and calendarEventId
# stays as the most recent one - the one the rest of the app treats as "where
# it currently is". Completing the to-do completes every one of those cards
# at once, since completion lives on the to-do, not the events.
def _normalize_todo_event_ids(todo):
    ids = todo.get('calendarEventIds')
    if not isinstance(ids, list):
        ids = [todo['calendarEventId']] if todo.get('calendarEventId') else []
        todo['calendarEventIds'] = ids
    return ids


# A to-do can sit on the calendar as several blocks ("splits"): the first one
# is the head, every other one points at it via splitOf, and splitPart/
# splitTotal ("2/3") are renumbered by start time whenever the group changes.
# Completion, title, details and attachments live on the to-do, so all blocks
# always mirror it.
def _event_split_group(events, event):
    head_id = event.get('splitOf') or event['id']
    group = [e for e in events if e['id'] == head_id or e.get('splitOf') == head_id]
    group.sort(key=lambda e: e.get('start') or '')
    return head_id, group


def _renumber_split_group(events, event):
    _, group = _event_split_group(events, event)
    if len(group) < 2:
        for e in group:
            e.pop('splitPart', None)
            e.pop('splitTotal', None)
        return
    for i, e in enumerate(group, 1):
        e['splitPart'] = i
        e['splitTotal'] = len(group)


def _todo_for_event(todos, event):
    for t in todos:
        ids = _normalize_todo_event_ids(t)
        if event['id'] in ids or (event.get('splitOf') and event['splitOf'] in ids):
            return t
    return None


def _sync_todo_to_events(todo, events):
    ids = set(_normalize_todo_event_ids(todo))
    now_iso = now_local().isoformat()
    for e in events:
        if e['id'] in ids or e.get('splitOf') in ids:
            e['title'] = todo['title']
            e['description'] = todo.get('details') or ''
            e['attachments'] = todo.get('attachments') or []
            e['updatedAt'] = now_iso


# A to-do whose latest scheduled day has fully passed (it's now a later
# calendar date) without being marked completed picks up a standing "Not
# Finished" flag - by design, this doesn't fire until the scheduled day is
# actually over (not just once its time slot ends), matching "not completed
# by 11:59:59pm". Its calendar card stays right where it is as a record of
# what was planned (and counts as one of the to-do's scheduled times - see
# calendarEventIds); the to-do just needs scheduling again, which adds a new
# card rather than replacing the old one. notFinished is derived from the
# latest card: past -> true, today or later (e.g. rescheduled or dragged
# forward) -> false. Runs on every load (called from load_todos, which every
# todos route goes through) rather than needing a scheduler, so it's always
# current regardless of which route a client happens to hit first.
def _apply_todo_rollovers(todos):
    today = now_local().date()
    events_by_id = None
    changed = False
    for todo in todos:
        had_ids = isinstance(todo.get('calendarEventIds'), list)
        ids = _normalize_todo_event_ids(todo)
        if not had_ids and ids:
            changed = True
        if todo.get('completed') or not todo.get('calendarEventId'):
            continue
        if events_by_id is None:
            events_by_id = {e['id']: e for e in load_calendar_events()}
        event = events_by_id.get(todo['calendarEventId'])
        if not event:
            continue
        try:
            event_date = datetime.fromisoformat(event.get('start') or '').date()
        except (ValueError, TypeError):
            continue
        missed = event_date < today
        if bool(todo.get('notFinished')) != missed:
            todo['notFinished'] = missed
            changed = True
    if changed:
        save_todos(todos)
    return todos


def save_todos(data):
    with open(TODOS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def load_todo_series():
    if os.path.exists(TODO_SERIES_FILE):
        try:
            with open(TODO_SERIES_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except (json.JSONDecodeError, OSError):
            pass
    return []


def save_todo_series(data):
    with open(TODO_SERIES_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def load_todo_templates():
    if os.path.exists(TODO_TEMPLATES_FILE):
        try:
            with open(TODO_TEMPLATES_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except (json.JSONDecodeError, OSError):
            pass
    return []


def save_todo_templates(data):
    with open(TODO_TEMPLATES_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


_RECURRENCE_FREQS = ('daily', 'weekly', 'monthly', 'yearly')


# Normalizes and sanity-checks a recurrence rule from the client into the
# shape _iterate_recurrence_dates expects. Only 'weekly' needs byWeekday,
# only 'monthly' needs byMonthday, only 'yearly' needs month/day - daily is
# just the interval. Returns (recurrence_dict, None) or (None, error_message).
def _validate_recurrence(raw):
    freq = str(raw.get('freq') or '').strip().lower()
    if freq not in _RECURRENCE_FREQS:
        return None, 'Please choose how often this should repeat.'
    try:
        interval = int(raw.get('interval') or 1)
    except (TypeError, ValueError):
        interval = 1
    if interval < 1:
        interval = 1
    try:
        start_date = date.fromisoformat(str(raw.get('start_date')))
    except (ValueError, TypeError):
        return None, 'Please choose a start date.'
    time_str = str(raw.get('time') or '').strip()
    try:
        _minutes_from_hhmm(time_str)
    except Exception:
        return None, 'Please choose a time of day.'
    end_date = None
    if raw.get('end_date'):
        try:
            end_date = date.fromisoformat(str(raw.get('end_date')))
        except (ValueError, TypeError):
            return None, 'Invalid end date.'
        if end_date < start_date:
            return None, 'The end date has to be after the start date.'

    recurrence = {
        'freq': freq, 'interval': interval, 'startDate': start_date.isoformat(),
        'time': time_str, 'endDate': end_date.isoformat() if end_date else None
    }

    if freq == 'weekly':
        try:
            by_weekday = sorted({int(d) for d in (raw.get('by_weekday') or []) if 0 <= int(d) <= 6})
        except (TypeError, ValueError):
            by_weekday = []
        if not by_weekday:
            by_weekday = [_day_of_week_sunday0(start_date)]
        recurrence['byWeekday'] = by_weekday
    elif freq == 'monthly':
        try:
            by_monthday = int(raw.get('by_monthday') or start_date.day)
        except (TypeError, ValueError):
            by_monthday = start_date.day
        recurrence['byMonthday'] = by_monthday if 1 <= by_monthday <= 31 else start_date.day
    elif freq == 'yearly':
        try:
            month = int(raw.get('month') or start_date.month)
            day = int(raw.get('day') or start_date.day)
        except (TypeError, ValueError):
            month, day = start_date.month, start_date.day
        recurrence['month'] = month
        recurrence['day'] = day

    return recurrence, None


# Every occurrence date a recurrence rule produces, strictly after
# `after_date` through `through_date` inclusive - this is exactly the kind of
# precise interval math that should never be left to a model to reason about
# (same reasoning as the bulk-schedule bin-packer above). Walks day by day
# rather than jumping by month/year directly - the window callers pass is
# always small (the 4-week rolling horizon), so the simpler, obviously-correct
# form costs nothing and there's no month/leap-year arithmetic to get subtly
# wrong.
def _iterate_recurrence_dates(recurrence, after_date, through_date):
    start_date = date.fromisoformat(recurrence['startDate'])
    if recurrence.get('endDate'):
        end_date = date.fromisoformat(recurrence['endDate'])
        if end_date < through_date:
            through_date = end_date
    freq = recurrence['freq']
    interval = recurrence['interval']
    start_week = start_date - timedelta(days=_day_of_week_sunday0(start_date))

    d = max(start_date, after_date + timedelta(days=1))
    while d <= through_date:
        matches = False
        if freq == 'daily':
            matches = (d - start_date).days % interval == 0
        elif freq == 'weekly':
            week_index = (d - start_week).days // 7
            matches = week_index % interval == 0 and _day_of_week_sunday0(d) in recurrence['byWeekday']
        elif freq == 'monthly':
            month_index = (d.year - start_date.year) * 12 + (d.month - start_date.month)
            matches = d.day == recurrence['byMonthday'] and month_index % interval == 0
        elif freq == 'yearly':
            matches = d.month == recurrence['month'] and d.day == recurrence['day'] and (d.year - start_date.year) % interval == 0
        if matches:
            yield d
        d += timedelta(days=1)


# Tops up every active series' occurrences to the rolling 4-week horizon (or
# its end date, if sooner) - called at the top of every /todos GET (see
# todos_list) so the window stays current on its own with no scheduler, same
# pattern as _apply_todo_rollovers. Each occurrence is created already
# scheduled (todo + linked calendar event, both real, both ordinary) at the
# series' fixed time of day - unlike Ashanti's bulk-schedule, there's no free-
# slot search here, since the whole point of a recurring to-do is the same
# time every time.
def _ensure_all_series_generated():
    series_list = load_todo_series()
    active = [s for s in series_list if s.get('active', True)]
    if not active:
        return
    horizon = today_local() + timedelta(days=28)
    if all(date.fromisoformat(s['lastGeneratedThrough']) >= horizon for s in active):
        return

    with todos_lock, calendar_lock, todo_series_lock:
        # Re-load under lock - series_list above was only read to decide
        # whether it's worth taking the locks at all.
        series_list = load_todo_series()
        todos = load_todos()
        working_events = load_calendar_events()
        now = now_local()
        changed = False

        for series in series_list:
            if not series.get('active', True):
                continue
            after = date.fromisoformat(series['lastGeneratedThrough'])
            if after >= horizon:
                continue
            for occ_date in _iterate_recurrence_dates(series['recurrence'], after, horizon):
                start_dt = datetime.combine(occ_date, datetime.min.time()) + timedelta(
                    minutes=_minutes_from_hhmm(series['recurrence']['time']))
                end_dt = start_dt + timedelta(minutes=series['estimatedMinutes'])
                now_iso = now.isoformat()
                event = {
                    'id': uuid.uuid4().hex, 'userId': series['userId'], 'title': series['title'],
                    'description': series.get('details') or '', 'location': '',
                    'start': start_dt.isoformat(), 'end': end_dt.isoformat(), 'allDay': False,
                    'source': 'internal', 'origin': 'todo', 'externalUid': None,
                    'status': 'confirmed', 'createdAt': now_iso, 'updatedAt': now_iso
                }
                working_events.append(event)
                todos.append({
                    'id': uuid.uuid4().hex, 'userId': series['userId'], 'title': series['title'],
                    'details': series.get('details') or '', 'estimatedMinutes': series['estimatedMinutes'],
                    'completed': False, 'createdAt': now_iso, 'completedAt': None,
                    'calendarEventId': event['id'], 'calendarEventIds': [event['id']], 'attachments': [],
                    'priority': series['priority'], 'personal': series.get('personal', False),
                    'notFinished': False, 'seriesId': series['id'], 'occurrenceDate': occ_date.isoformat()
                })
                changed = True
            series['lastGeneratedThrough'] = horizon.isoformat()
            changed = True

        if changed:
            save_calendar_events(working_events)
            save_todos(todos)
            save_todo_series(series_list)


@app.route('/todos', methods=['GET'])
def todos_list():
    _ensure_lunch_generated()
    _reevaluate_upcoming_lunches()
    _ensure_all_series_generated()
    return jsonify({'success': True, 'todos': load_todos()})


@app.route('/todos', methods=['POST'])
def todos_create():
    try:
        data = request.json or {}
        title = str(data.get('title', '')).strip()
        details = str(data.get('details', '') or '').strip()
        try:
            estimated_minutes = int(data.get('estimated_minutes'))
        except (TypeError, ValueError):
            estimated_minutes = 0
        if not title or not details or estimated_minutes <= 0:
            return jsonify({'success': False, 'error': 'Missing title, details, or estimated time'}), 400
        priority = str(data.get('priority') or 'medium').strip().lower()
        if priority not in ('high', 'medium', 'low'):
            priority = 'medium'
        with todos_lock:
            todos = load_todos()
            todo = {
                'id': uuid.uuid4().hex,
                'userId': current_user_id(),
                'title': title,
                'details': details,
                'estimatedMinutes': estimated_minutes,
                'completed': False,
                'createdAt': now_local().isoformat(),
                'completedAt': None,
                'calendarEventId': None,
                'calendarEventIds': [],
                'attachments': data.get('attachments') or [],
                # High: must land in the period it's scheduled for. Medium:
                # should land there, but rolls to the next equivalent period
                # if it doesn't fit. Low: pure filler, scheduled only with
                # whatever room is left over - see /todos/bulk-schedule.
                'priority': priority,
                # Personal todos only ever get scheduled outside business
                # hours (see compute_personal_free_slots) - a business todo
                # never lands there and vice versa. Its calendar card's color
                # comes from the single global personal-to-do color setting
                # (see load_todo_settings), not anything stored per to-do.
                'personal': bool(data.get('personal')),
                # Set by _apply_todo_rollovers when a scheduled day passes
                # with the to-do still incomplete - see that function.
                'notFinished': False
            }
            todos.append(todo)
            save_todos(todos)
        return jsonify({'success': True, 'todos': todos, 'todo_id': todo['id']})
    except Exception as e:
        print(f"Todo create error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/todos/<todo_id>', methods=['POST'])
def todos_update(todo_id):
    try:
        data = request.json or {}
        with todos_lock, calendar_lock:
            todos = load_todos()
            todo = next((t for t in todos if t['id'] == todo_id), None)
            if not todo:
                return jsonify({'success': False, 'error': 'To-do not found'}), 404
            if 'title' in data:
                new_title = str(data['title'] or '').strip()
                if new_title:
                    todo['title'] = new_title
            if 'details' in data:
                todo['details'] = str(data['details'] or '').strip()
            if 'estimated_minutes' in data:
                try:
                    new_minutes = int(data['estimated_minutes'])
                    if new_minutes > 0:
                        todo['estimatedMinutes'] = new_minutes
                except (TypeError, ValueError):
                    pass
            if 'completed' in data:
                todo['completed'] = bool(data['completed'])
                todo['completedAt'] = now_local().isoformat() if todo['completed'] else None
                # Resolves the "Not Finished" flag the same as a fresh
                # reschedule does - it means "missed at least one window",
                # not a permanent mark once the to-do is actually done.
                if todo['completed']:
                    todo['notFinished'] = False
            if 'attachments' in data:
                todo['attachments'] = data['attachments'] or []
            if 'priority' in data:
                new_priority = str(data['priority'] or '').strip().lower()
                if new_priority in ('high', 'medium', 'low'):
                    todo['priority'] = new_priority
            if 'personal' in data:
                todo['personal'] = bool(data['personal'])
            # A new id is another time this to-do is scheduled - it joins
            # calendarEventIds (earlier cards, e.g. a missed one, stay) and
            # becomes the current calendarEventId. null explicitly clears
            # every scheduling, the start-from-scratch "Unschedule".
            if 'calendar_event_id' in data:
                new_event_id = data['calendar_event_id'] or None
                ids = _normalize_todo_event_ids(todo)
                if new_event_id:
                    if new_event_id not in ids:
                        ids.append(new_event_id)
                    todo['calendarEventId'] = new_event_id
                    # A fresh scheduling attempt resolves "Not Finished" -
                    # it only means the latest card was missed.
                    todo['notFinished'] = False
                else:
                    todo['calendarEventId'] = None
                    todo['calendarEventIds'] = []
                    todo['notFinished'] = False
            # Every block of a split to-do mirrors its text - an edit here
            # reaches all of them.
            if any(k in data for k in ('title', 'details', 'attachments')):
                events = load_calendar_events()
                _sync_todo_to_events(todo, events)
                save_calendar_events(events)
            save_todos(todos)
        return jsonify({'success': True, 'todos': todos})
    except Exception as e:
        print(f"Todo update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/todos/<todo_id>', methods=['DELETE'])
def todos_delete(todo_id):
    try:
        with todos_lock:
            todos = load_todos()
            todo = next((t for t in todos if t['id'] == todo_id), None)
            if not todo:
                return jsonify({'success': False, 'error': 'To-do not found'}), 404
            todos = [t for t in todos if t['id'] != todo_id]
            save_todos(todos)
        _delete_item_attachments(todo)
        return jsonify({'success': True, 'todos': todos})
    except Exception as e:
        print(f"Todo delete error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# --- Recurring to-dos -------------------------------------------------------
# A series is the template + rule a recurring to-do was set up with - see
# _ensure_all_series_generated for how it actually spawns occurrences.
@app.route('/todos/series', methods=['POST'])
def todos_series_create():
    try:
        data = request.json or {}
        title = str(data.get('title', '')).strip()
        details = str(data.get('details', '') or '').strip()
        try:
            estimated_minutes = int(data.get('estimated_minutes'))
        except (TypeError, ValueError):
            estimated_minutes = 0
        if not title or not details or estimated_minutes <= 0:
            return jsonify({'success': False, 'error': 'Missing title, details, or estimated time'}), 400
        priority = str(data.get('priority') or 'medium').strip().lower()
        if priority not in ('high', 'medium', 'low'):
            priority = 'medium'

        recurrence, err = _validate_recurrence(data.get('recurrence') or {})
        if err:
            return jsonify({'success': False, 'error': err}), 400

        series = {
            'id': uuid.uuid4().hex, 'userId': current_user_id(), 'title': title,
            'details': details, 'estimatedMinutes': estimated_minutes, 'priority': priority,
            'personal': bool(data.get('personal')),
            'recurrence': recurrence, 'active': True,
            'createdAt': now_local().isoformat(),
            # One day before its own first occurrence, so the very first
            # generation pass (triggered right below) picks that date up too.
            'lastGeneratedThrough': (date.fromisoformat(recurrence['startDate']) - timedelta(days=1)).isoformat()
        }
        with todo_series_lock:
            series_list = load_todo_series()
            series_list.append(series)
            save_todo_series(series_list)

        _ensure_all_series_generated()

        return jsonify({'success': True, 'series': series, 'todos': load_todos()})
    except Exception as e:
        print(f"Todo series create error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Edits the series template - used when a "this and all future" edit is made
# from either the To-Do page or a linked calendar card (see saveEventModal /
# renderTodosPreviewPane on the frontend), so occurrences generated *after*
# this point pick up the change too. The frontend applies the same fields to
# each already-existing future occurrence itself via the ordinary
# /todos/<id> and /calendar/events/<id> routes - this route only ever touches
# the series definition, never an occurrence.
@app.route('/todos/series/<series_id>', methods=['POST'])
def todos_series_update(series_id):
    try:
        data = request.json or {}
        with todo_series_lock:
            series_list = load_todo_series()
            series = next((s for s in series_list if s['id'] == series_id), None)
            if not series:
                return jsonify({'success': False, 'error': 'Series not found'}), 404
            if 'title' in data:
                new_title = str(data['title'] or '').strip()
                if new_title:
                    series['title'] = new_title
            if 'details' in data:
                series['details'] = str(data['details'] or '').strip()
            if 'estimated_minutes' in data:
                try:
                    new_minutes = int(data['estimated_minutes'])
                    if new_minutes > 0:
                        series['estimatedMinutes'] = new_minutes
                except (TypeError, ValueError):
                    pass
            if 'priority' in data:
                new_priority = str(data['priority'] or '').strip().lower()
                if new_priority in ('high', 'medium', 'low'):
                    series['priority'] = new_priority
            if 'personal' in data:
                series['personal'] = bool(data['personal'])
            # Only the time of day is ever edited after the series exists
            # (dragging/resizing a linked event and choosing "all future") -
            # the frequency/weekday/month-day rule itself isn't editable in
            # place, only by ending this series and starting a new one.
            if 'time' in data:
                try:
                    _minutes_from_hhmm(data['time'])
                    series['recurrence']['time'] = data['time']
                except Exception:
                    pass
            save_todo_series(series_list)
        return jsonify({'success': True, 'series': series})
    except Exception as e:
        print(f"Todo series update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Ends a series (no more occurrences ever get generated past this point) and,
# unless told to keep them, removes every occurrence that's today or later
# and not already completed - completed and past occurrences are left alone
# as history, matching the same "future only" scope an "apply to all" edit
# uses (see _ensure_all_series_generated's comment and saveEventModal).
@app.route('/todos/series/<series_id>', methods=['DELETE'])
def todos_series_delete(series_id):
    try:
        delete_future = str(request.args.get('scope', 'future')).lower() != 'keep'
        with todo_series_lock:
            series_list = load_todo_series()
            series = next((s for s in series_list if s['id'] == series_id), None)
            if not series:
                return jsonify({'success': False, 'error': 'Series not found'}), 404
            series['active'] = False
            save_todo_series(series_list)

        removed_todo_ids = []
        if delete_future:
            today_iso = today_local().isoformat()
            with todos_lock, calendar_lock:
                todos = load_todos()
                to_remove = [
                    t for t in todos
                    if t.get('seriesId') == series_id and not t.get('completed')
                    and (t.get('occurrenceDate') or '') >= today_iso
                ]
                if to_remove:
                    removed_todo_ids = [t['id'] for t in to_remove]
                    event_ids = {eid for t in to_remove for eid in _normalize_todo_event_ids(t)}
                    event_ids |= {t['calendarEventId'] for t in to_remove if t.get('calendarEventId')}
                    todos = [t for t in todos if t['id'] not in removed_todo_ids]
                    working_events = [e for e in load_calendar_events() if e['id'] not in event_ids]
                    save_calendar_events(working_events)
                    save_todos(todos)
                    for t in to_remove:
                        _delete_item_attachments(t)

        return jsonify({'success': True, 'todos': load_todos(), 'removedTodoIds': removed_todo_ids})
    except Exception as e:
        print(f"Todo series delete error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/todos/templates', methods=['GET'])
def todo_templates_list():
    return jsonify({'success': True, 'templates': load_todo_templates()})


@app.route('/todos/templates', methods=['POST'])
def todo_templates_create():
    try:
        data = request.json or {}
        title = str(data.get('title', '')).strip()
        details = str(data.get('details', '') or '').strip()
        try:
            estimated_minutes = int(data.get('estimated_minutes'))
        except (TypeError, ValueError):
            estimated_minutes = 0
        if not title:
            return jsonify({'success': False, 'error': 'Please name the template.'}), 400
        priority = str(data.get('priority') or 'medium').strip().lower()
        if priority not in ('high', 'medium', 'low'):
            priority = 'medium'
        template = {
            'id': uuid.uuid4().hex,
            'title': title,
            'details': details,
            'estimatedMinutes': estimated_minutes if estimated_minutes > 0 else 0,
            'priority': priority,
            'personal': bool(data.get('personal')),
            'createdAt': now_local().isoformat()
        }
        with todo_templates_lock:
            templates = load_todo_templates()
            templates.append(template)
            save_todo_templates(templates)
        return jsonify({'success': True, 'templates': templates})
    except Exception as e:
        print(f"Todo template create error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/todos/templates/<template_id>', methods=['POST'])
def todo_templates_update(template_id):
    try:
        data = request.json or {}
        with todo_templates_lock:
            templates = load_todo_templates()
            template = next((t for t in templates if t['id'] == template_id), None)
            if not template:
                return jsonify({'success': False, 'error': 'Template not found'}), 404
            if 'title' in data:
                new_title = str(data['title'] or '').strip()
                if new_title:
                    template['title'] = new_title
            if 'details' in data:
                template['details'] = str(data['details'] or '').strip()
            if 'estimated_minutes' in data:
                try:
                    template['estimatedMinutes'] = max(0, int(data['estimated_minutes']))
                except (TypeError, ValueError):
                    pass
            if 'priority' in data:
                new_priority = str(data['priority'] or '').strip().lower()
                if new_priority in ('high', 'medium', 'low'):
                    template['priority'] = new_priority
            if 'personal' in data:
                template['personal'] = bool(data['personal'])
            save_todo_templates(templates)
        return jsonify({'success': True, 'templates': templates})
    except Exception as e:
        print(f"Todo template update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/todos/templates/<template_id>', methods=['DELETE'])
def todo_templates_delete(template_id):
    try:
        with todo_templates_lock:
            templates = load_todo_templates()
            if not any(t['id'] == template_id for t in templates):
                return jsonify({'success': False, 'error': 'Template not found'}), 404
            templates = [t for t in templates if t['id'] != template_id]
            save_todo_templates(templates)
        return jsonify({'success': True, 'templates': templates})
    except Exception as e:
        print(f"Todo template delete error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Deterministic bin-packing scheduler behind the To-Do bulk-schedule action -
# see the priority rules in the create_project_task/create_file style comment
# blocks elsewhere: this is exactly the kind of precise interval math that
# should never be left to a model to reason about on the fly (same lesson as
# the 15-minute post-meeting buffer above) - Ashanti's own reply just narrates
# what this function actually did.
#
# Each to-do carries its own period, chosen per-row in the modal (see
# buildBulkScheduleTodoRow) - that's the one Francis picked, so nothing here
# ever silently reschedules it to a different period. High gets first pick of
# its own period's free time, then Medium, then Low (same contested day = same
# priority order as before), but each to-do gets exactly one shot at exactly
# the period it was given. Anything that doesn't fit comes back unscheduled
# with a reason, for Francis to reschedule to a different period - see
# previewBulkScheduleTodos.
#
# `preview` (bool, default False) runs the exact same slot-finding pass but
# skips creating/saving anything - the bulk-schedule popup uses this so
# clicking "Schedule" only places to-dos on its live preview calendar
# (bulkScheduleStagedPlacements), nothing becomes real until Save actually
# commits what's still on that calendar (see saveBulkScheduleTodos). A
# preview response carries each placement's start/end instead of a saved
# event, since there is no event yet.
@app.route('/todos/bulk-schedule', methods=['POST'])
def todos_bulk_schedule():
    try:
        # Lunch has to already be sitting on the calendar as a real event
        # before this does its own free-slot search, or it would happily
        # bin-pack a to-do right into that window - see _ensure_lunch_generated.
        # Reverting any now-unnecessary shift first (see
        # _reevaluate_upcoming_lunches) means this search starts from
        # lunch's true minimal footprint rather than one still reflecting a
        # to-do that's since been removed.
        _ensure_lunch_generated()
        _reevaluate_upcoming_lunches()
        data = request.json or {}
        preview = bool(data.get('preview'))
        raw_items = data.get('items') or []
        reserved_raw = data.get('reserved') or []
        if not raw_items:
            return jsonify({'success': False, 'error': 'No to-dos selected'}), 400

        items = []
        for raw in raw_items:
            todo_id = raw.get('todo_id')
            try:
                start_date = date.fromisoformat(str(raw.get('start_date')))
                end_date = date.fromisoformat(str(raw.get('end_date')))
            except (ValueError, TypeError):
                return jsonify({'success': False, 'error': 'Invalid start_date/end_date'}), 400
            if not todo_id or end_date < start_date:
                return jsonify({'success': False, 'error': 'Invalid to-do or date range'}), 400
            items.append({'todo_id': todo_id, 'start_date': start_date, 'end_date': end_date})

        now = now_local()

        with todos_lock, calendar_lock:
            todos = load_todos()
            todos_by_id = {t['id']: t for t in todos}
            valid_items = [it for it in items if it['todo_id'] in todos_by_id and not todos_by_id[it['todo_id']].get('completed')]
            if not valid_items:
                return jsonify({'success': False, 'error': 'None of the selected to-dos were found'}), 400

            working_events = load_calendar_events()

            # A completed to-do's calendar time is free again - the work is done, so
            # its card (which stays on the calendar as a record) shouldn't block
            # anything from being scheduled there. Only hidden from the free-slot
            # search below; the cards themselves are never touched or removed.
            freed_event_ids = set()
            for t in todos:
                if t.get('completed'):
                    freed_event_ids.update(_normalize_todo_event_ids(t))
            freed_event_ids |= {e.get('id') for e in working_events if e.get('splitOf') in freed_event_ids}

            # Placements the popup already has staged locally but hasn't
            # saved yet (see bulkScheduleStagedPlacements/previewBulkScheduleTodos)
            # - without these, scheduling a second batch in the same popup
            # session has no way to know what an earlier batch already
            # claimed and happily lands right on top of it, since nothing
            # staged is a real, saved event yet for this search to see.
            # Synthetic and never persisted - filtered back out below,
            # before anything is saved.
            for r in reserved_raw:
                try:
                    working_events.append({
                        'start': str(r['start']), 'end': str(r['end']),
                        'allDay': False, 'status': 'confirmed', '_synthetic': True
                    })
                except (KeyError, TypeError):
                    continue

            def day_events_for(day):
                day_iso = day.isoformat()
                return [e for e in working_events
                        if e.get('status') != 'cancelled' and e.get('id') not in freed_event_ids
                        and (e.get('start') or '').startswith(day_iso)]

            # If a plain search finds no room for `duration` on `day`, and
            # that day has a real Lunch card on it, tries repositioning
            # lunch to either extreme of its own flexibility window (the
            # only two positions that can ever open up MORE contiguous room
            # than it already has - anything between them just trades room
            # from one side of the day to the other) to see whether that
            # alone would make the to-do fit. Only actually commits the move
            # if a candidate both fits the to-do AND doesn't land lunch on
            # top of anything else real that day - lunch has to never be
            # skipped and never overlap anything, so a shift that would
            # violate the second guarantee to satisfy the first is rejected
            # outright, same as a candidate that doesn't help at all.
            moved_lunch_events = {}

            def try_shift_lunch_for_slot(day, duration, slot_fn):
                day_iso = day.isoformat()
                lunch_event = next(
                    (e for e in working_events if e.get('origin') == 'lunch'
                     and e.get('status') != 'cancelled' and (e.get('start') or '').startswith(day_iso)),
                    None
                )
                if not lunch_event:
                    return None
                lunch_settings = load_lunch_settings()
                if not lunch_settings['days']:
                    return None
                try:
                    original_start = lunch_event['start']
                    original_end = lunch_event['end']
                    orig_start_dt = datetime.fromisoformat(original_start)
                except (ValueError, KeyError, TypeError):
                    return None

                length_min = lunch_settings['lengthMinutes']
                natural_start_min = _minutes_from_hhmm(lunch_settings['startTime'])
                window_start = max(0, natural_start_min - lunch_settings['flexBeforeMinutes'])
                window_end = min(24 * 60, natural_start_min + length_min + lunch_settings['flexAfterMinutes'])
                day_start = datetime.combine(day, datetime.min.time())
                orig_start_min = (orig_start_dt - day_start).total_seconds() / 60

                other_events = [e for e in day_events_for(day) if e is not lunch_event]

                def overlaps_others(start_dt, end_dt):
                    for e in other_events:
                        if e.get('allDay'):
                            continue
                        try:
                            es = datetime.fromisoformat(e['start'])
                            ee = datetime.fromisoformat(e.get('end') or e['start'])
                        except (ValueError, KeyError, TypeError):
                            continue
                        if es < end_dt and start_dt < ee:
                            return True
                    return False

                for candidate_min in (window_start, window_end - length_min):
                    if candidate_min == orig_start_min:
                        continue
                    candidate_start = day_start + timedelta(minutes=candidate_min)
                    candidate_end = candidate_start + timedelta(minutes=length_min)
                    if overlaps_others(candidate_start, candidate_end):
                        continue
                    lunch_event['start'] = candidate_start.isoformat()
                    lunch_event['end'] = candidate_end.isoformat()
                    for s, en in slot_fn(day_events_for(day), day, now):
                        if en - s >= duration:
                            lunch_event['updatedAt'] = now.isoformat()
                            moved_lunch_events[lunch_event['id']] = lunch_event
                            return (s, s + duration)
                    lunch_event['start'] = original_start
                    lunch_event['end'] = original_end

                return None

            # Only tried for a to-do of 4 hours (240 min) or longer, and only
            # once a plain search AND a lunch shift have both already failed
            # for `day` - splitting is the last resort, not the first idea.
            # Finds the free stretch immediately before lunch and the one
            # immediately after it, and - if each is at least an hour - divides
            # the to-do's full duration between them (favoring filling the
            # side with more natural room first, falling back to the other
            # side if that leaves too little for the far side), so it reads
            # as "as much as fits before lunch, the rest right after" rather
            # than an arbitrary 50/50 cut. Returns None if either side would
            # end up under an hour.
            def try_split_around_lunch(day, duration_min, slot_fn):
                day_iso = day.isoformat()
                lunch_event = next(
                    (e for e in working_events if e.get('origin') == 'lunch'
                     and e.get('status') != 'cancelled' and (e.get('start') or '').startswith(day_iso)),
                    None
                )
                if not lunch_event:
                    return None
                try:
                    lunch_start = datetime.fromisoformat(lunch_event['start'])
                    lunch_end = datetime.fromisoformat(lunch_event['end'])
                except (ValueError, KeyError, TypeError):
                    return None

                free = slot_fn(day_events_for(day), day, now)
                before = [(s, e) for s, e in free if e <= lunch_start]
                after = [(s, e) for s, e in free if s >= lunch_end]
                if not before or not after:
                    return None
                b_s, b_e = before[-1]
                a_s, a_e = after[0]
                avail_before = (b_e - b_s).total_seconds() / 60
                avail_after = (a_e - a_s).total_seconds() / 60
                if avail_before < 60 or avail_after < 60:
                    return None

                piece1 = min(avail_before, duration_min - 60)
                piece2 = duration_min - piece1
                if piece2 < 60 or piece2 > avail_after:
                    piece2 = min(avail_after, duration_min - 60)
                    piece1 = duration_min - piece2
                    if piece1 < 60 or piece1 > avail_before:
                        return None

                part1_end = b_e
                part1_start = part1_end - timedelta(minutes=piece1)
                part2_start = a_s
                part2_end = part2_start + timedelta(minutes=piece2)
                return (part1_start, part1_end, part2_start, part2_end)

            def find_slot(todo, days):
                duration_min = todo.get('estimatedMinutes') or 30
                duration = timedelta(minutes=duration_min)
                slot_fn = compute_personal_free_slots if todo.get('personal') else compute_free_slots
                for day in days:
                    for s, en in slot_fn(day_events_for(day), day, now):
                        if en - s >= duration:
                            return [(s, s + duration)]
                    shifted = try_shift_lunch_for_slot(day, duration, slot_fn)
                    if shifted:
                        return [shifted]
                    if duration_min >= 240:
                        split = try_split_around_lunch(day, duration_min, slot_fn)
                        if split:
                            return [(split[0], split[1]), (split[2], split[3])]
                return None

            def create_event_for(todo, start_dt, end_dt, split_part=None, split_of=None):
                now_iso = now.isoformat()
                event = {
                    'id': uuid.uuid4().hex,
                    'userId': current_user_id(),
                    'title': todo['title'],
                    'description': todo.get('details') or '',
                    'location': '',
                    'start': start_dt.isoformat(),
                    'end': end_dt.isoformat(),
                    'allDay': False,
                    'source': 'internal',
                    'origin': 'todo',
                    'externalUid': None,
                    'status': 'confirmed',
                    'createdAt': now_iso,
                    'updatedAt': now_iso
                }
                # A to-do split around lunch (see try_split_around_lunch)
                # gets two ordinary events rather than a new kind of record -
                # every existing single-event code path (drag, resize, edit,
                # complete, delete) keeps working on each half unchanged.
                # splitPart/splitTotal are display-only ("1 of 2"); splitOf on
                # the second half is what findTodoForCalendarEvent (index.html)
                # actually follows back to the to-do, since only the first
                # half is ever written to calendarEventId.
                if split_part:
                    event['splitPart'] = split_part
                    event['splitTotal'] = 2
                    if split_of:
                        event['splitOf'] = split_of
                working_events.append(event)
                return event

            # High priority across the whole batch gets first pick of its own
            # period's free time, then Medium, then Low - same ordering as
            # before, just no longer sharing one period to pick from.
            priority_rank = {'high': 0, 'medium': 1, 'low': 2}
            valid_items.sort(key=lambda it: priority_rank.get(todos_by_id[it['todo_id']].get('priority', 'medium'), 1))

            scheduled = []
            unscheduled = []

            for it in valid_items:
                todo = todos_by_id[it['todo_id']]
                priority = todo.get('priority', 'medium')
                period_days = [it['start_date'] + timedelta(days=i) for i in range((it['end_date'] - it['start_date']).days + 1)]

                parts = find_slot(todo, period_days)

                if parts:
                    # Always append to working_events, preview or not - a
                    # later item in this same batch has to see this slot as
                    # taken (day_events_for reads from working_events), or
                    # everything sharing a period piles up on the same first
                    # open slot instead of filling the period in order. Only
                    # whether the batch is actually persisted (below) depends
                    # on preview.
                    if len(parts) == 2:
                        event1 = create_event_for(todo, *parts[0], split_part=1)
                        event2 = create_event_for(todo, *parts[1], split_part=2, split_of=event1['id'])
                        events = [event1, event2]
                    else:
                        events = [create_event_for(todo, *parts[0])]
                    scheduled.append((todo, parts, events))
                elif priority == 'high':
                    unscheduled.append((todo, "No availability in the selected period, even at highest priority - try a different period."))
                elif priority == 'medium':
                    unscheduled.append((todo, "No availability in the selected period - try a different period."))
                else:
                    unscheduled.append((todo, "No free time left over in the selected period - try a different period."))

            if not preview:
                save_calendar_events([e for e in working_events if not e.get('_synthetic')])
                for todo, parts, events in scheduled:
                    todo['calendarEventId'] = events[0]['id']
                    # Same as a manual reschedule (see todos_update) - this
                    # writes the link directly rather than going through
                    # that route, so it has to join calendarEventIds and
                    # clear notFinished itself too.
                    ids = _normalize_todo_event_ids(todo)
                    if events[0]['id'] not in ids:
                        ids.append(events[0]['id'])
                    todo['notFinished'] = False
                save_todos(todos)

        return jsonify({
            'success': True,
            'scheduled': [
                {
                    'todoId': todo['id'], 'title': todo['title'],
                    'start': parts[0][0].isoformat(), 'end': parts[-1][1].isoformat(),
                    'parts': [{'start': s.isoformat(), 'end': en.isoformat()} for s, en in parts],
                    # A preview event's id is never persisted (see above) -
                    # leaving it out of the response keeps a discarded id
                    # from ever reaching the client.
                    **({'events': events} if not preview else {})
                }
                for todo, parts, events in scheduled
            ],
            'unscheduled': [
                {'todoId': todo['id'], 'title': todo['title'], 'reason': reason}
                for todo, reason in unscheduled
            ],
            # Any Lunch card this call repositioned to make room for a to-do
            # (see try_shift_lunch_for_slot) - sent even in preview mode so
            # the popup's own preview calendar can show where lunch actually
            # ends up, not just where it used to be.
            'lunchMoved': [
                {'id': e['id'], 'start': e['start'], 'end': e['end']}
                for e in moved_lunch_events.values()
            ]
        })
    except Exception as e:
        print(f"Bulk schedule error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Surfaces once per day, per to-do - same shown-status pattern as
# calendar_checkin_status.json (see /calendar/checkins/today), so a to-do
# already asked about today doesn't get re-asked on the next poll. A to-do
# Francis never acted on (didn't mark done or reschedule) is still
# notFinished tomorrow, so it comes right back on the next day's first poll -
# this isn't a one-time nag, it's meant to keep coming back until resolved.
@app.route('/todos/morning-checkin', methods=['GET'])
def todos_morning_checkin():
    try:
        today_iso = today_local().isoformat()
        status = load_todo_checkin_status()
        status = {k: v for k, v in status.items() if k >= (today_local() - timedelta(days=2)).isoformat()}
        shown_today = set(status.get(today_iso, []))

        pending = [
            t for t in load_todos()
            if t.get('notFinished') and not t.get('completed') and t['id'] not in shown_today
        ]

        if not pending:
            save_todo_checkin_status(status)
            return jsonify({'success': True, 'checkin': None})

        status[today_iso] = list(shown_today | {t['id'] for t in pending})
        save_todo_checkin_status(status)

        return jsonify({
            'success': True,
            'checkin': {'todos': [{'id': t['id'], 'title': t['title']} for t in pending]}
        })
    except Exception as ex:
        print(f"Todo morning checkin error: {ex}")
        return jsonify({'success': False, 'error': str(ex)}), 500


# --- Suggestions -----------------------------------------------------------
# A flat, standalone feature-request backlog - same shape as To-Do (title +
# details + completed), but deliberately has zero ties to agents, the
# calendar, or the Workspace. Just a personal punch list of things to build
# into the app later.
SUGGESTIONS_FILE = _data_path('suggestions.json')
suggestions_lock = threading.Lock()


def load_suggestions():
    if os.path.exists(SUGGESTIONS_FILE):
        try:
            with open(SUGGESTIONS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []
    return []


def save_suggestions(suggestions):
    with open(SUGGESTIONS_FILE, 'w', encoding='utf-8') as f:
        json.dump(suggestions, f, indent=2)


@app.route('/suggestions', methods=['GET'])
def suggestions_list():
    return jsonify({'success': True, 'suggestions': load_suggestions()})


@app.route('/suggestions', methods=['POST'])
def suggestions_create():
    try:
        data = request.json or {}
        title = str(data.get('title', '')).strip()
        details = str(data.get('details', '') or '').strip()
        if not title or not details:
            return jsonify({'success': False, 'error': 'Missing title or details'}), 400
        with suggestions_lock:
            suggestions = load_suggestions()
            suggestion = {
                'id': uuid.uuid4().hex,
                'userId': current_user_id(),
                'title': title,
                'details': details,
                'completed': False,
                'createdAt': now_local().isoformat(),
                'completedAt': None
            }
            suggestions.append(suggestion)
            save_suggestions(suggestions)
        return jsonify({'success': True, 'suggestions': suggestions, 'suggestion_id': suggestion['id']})
    except Exception as e:
        print(f"Suggestions create error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/suggestions/<suggestion_id>', methods=['POST'])
def suggestions_update(suggestion_id):
    try:
        data = request.json or {}
        with suggestions_lock:
            suggestions = load_suggestions()
            suggestion = next((s for s in suggestions if s['id'] == suggestion_id), None)
            if not suggestion:
                return jsonify({'success': False, 'error': 'Suggestion not found'}), 404
            if 'title' in data:
                new_title = str(data['title'] or '').strip()
                if new_title:
                    suggestion['title'] = new_title
            if 'details' in data:
                suggestion['details'] = str(data['details'] or '').strip()
            if 'completed' in data:
                suggestion['completed'] = bool(data['completed'])
                suggestion['completedAt'] = now_local().isoformat() if suggestion['completed'] else None
            save_suggestions(suggestions)
        return jsonify({'success': True, 'suggestions': suggestions})
    except Exception as e:
        print(f"Suggestions update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/suggestions/<suggestion_id>', methods=['DELETE'])
def suggestions_delete(suggestion_id):
    try:
        with suggestions_lock:
            suggestions = load_suggestions()
            if not any(s['id'] == suggestion_id for s in suggestions):
                return jsonify({'success': False, 'error': 'Suggestion not found'}), 404
            suggestions = [s for s in suggestions if s['id'] != suggestion_id]
            save_suggestions(suggestions)
        return jsonify({'success': True, 'suggestions': suggestions})
    except Exception as e:
        print(f"Suggestions delete error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Reorders the whole list to match `order` (a list of ids, active items
# only - see moveSuggestion in index.html, which only offers up/down within
# the active section since Completed is already sorted by completion date).
# Anything not mentioned in `order` is kept, appended at the end, rather
# than silently dropped - defensive against a stale client-side list.
@app.route('/suggestions/reorder', methods=['POST'])
def suggestions_reorder():
    try:
        data = request.json or {}
        order = data.get('order') or []
        with suggestions_lock:
            suggestions = load_suggestions()
            by_id = {s['id']: s for s in suggestions}
            reordered = [by_id[i] for i in order if i in by_id]
            seen = set(order)
            missing = [s for s in suggestions if s['id'] not in seen]
            suggestions = reordered + missing
            save_suggestions(suggestions)
        return jsonify({'success': True, 'suggestions': suggestions})
    except Exception as e:
        print(f"Suggestions reorder error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# --- Lana: language learning ------------------------------------------------
# Settings (Settings > Agents > Lana) hold the languages Francis wants to learn
# and how long he can study each day. Lana turns a language into a PROJECT: each
# lesson is a task in it, prepared as an interactive lesson (create_lesson) that
# opens in the Workspace, where he answers the exercises himself. When he
# finishes a lesson its results go back to Lana, who can add more lessons to the
# project (add_lessons) depending on how it went.
LANA_SETTINGS_FILE = _data_path('lana_settings.json')
lana_settings_lock = threading.Lock()
LANA_DAILY_MINUTES = (15, 30, 45, 60, 90, 120)


def load_lana_settings():
    data = {}
    if os.path.exists(LANA_SETTINGS_FILE):
        try:
            with open(LANA_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                data = loaded
        except (json.JSONDecodeError, OSError):
            pass
    languages = [l for l in (data.get('languages') or []) if isinstance(l, dict) and str(l.get('name') or '').strip()]
    minutes = data.get('dailyMinutes')
    return {'languages': languages, 'dailyMinutes': minutes if minutes in LANA_DAILY_MINUTES else 30}


def save_lana_settings(data):
    with open(LANA_SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


@app.route('/lana/settings', methods=['GET'])
def lana_settings_get():
    return jsonify({'success': True, 'settings': load_lana_settings()})


@app.route('/lana/settings', methods=['POST'])
def lana_settings_save():
    try:
        data = request.json or {}
        languages = []
        seen = set()
        for raw in (data.get('languages') or [])[:8]:
            if not isinstance(raw, dict):
                continue
            name = str(raw.get('name') or '').strip()[:40]
            if not name or name.lower() in seen:
                continue
            seen.add(name.lower())
            languages.append({
                'id': re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-') or uuid.uuid4().hex[:8],
                'name': name,
                'focus': str(raw.get('focus') or '').strip()[:120]
            })
        try:
            minutes = int(data.get('daily_minutes'))
        except (TypeError, ValueError):
            minutes = 30
        if minutes not in LANA_DAILY_MINUTES:
            minutes = 30
        settings = {'languages': languages, 'dailyMinutes': minutes}
        with lana_settings_lock:
            save_lana_settings(settings)
        return jsonify({'success': True, 'settings': settings})
    except Exception as e:
        print(f"Lana settings error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# What Lana is told about the languages Francis picked, every turn.
def get_lana_context():
    settings = load_lana_settings()
    languages = settings['languages']
    if languages:
        listing = "; ".join(l['name'] + (f" ({l['focus']})" if l.get('focus') else '') for l in languages)
        lang_line = f"Languages Francis chose to learn in Settings: {listing}."
    else:
        lang_line = "Francis hasn't picked a language to learn in Settings yet - if he wants to start one, ask which."
    return (
        "\n\nLANGUAGE SETTINGS (set by Francis in Settings > Agents > Lana): "
        f"{lang_line} He can study about {settings['dailyMinutes']} minutes a day. "
        "Use these instead of asking again. Only ask what you still need (goal, preferred style, level if he isn't a "
        "complete beginner), one question at a time."
    )


LANA_LESSON_RULES = (
    "\n\nHow Lana's course works: a language course is an INTERACTIVE PROJECT - a folder you and Francis work through together (modules, lessons, exercises, quizzes) - not a job with tasks. The goal is real: take him "
    "from NOVICE to PROFICIENT, so the plan is long. Lay out the whole journey in propose_language_plan as a syllabus: "
    "MODULES grouped by level (A1 beginner, A2, B1, B2, C1 - about 8 to 12 modules in all, each with a CEFR level and a "
    "rough timeframe). The FIRST module is spelled out as numbered LESSONS (about 12 to 15, sized to his daily study time), "
    "each with its specific TOPICS (the actual words, forms and skills it teaches - concrete, like \"all six forms of "
    "avere\"). Every later module is a ROADMAP: what it will cover and roughly how many lessons - you turn it into real "
    "lessons once he finishes the module before it, based on how he did. Don't number lessons in their names (the app "
    "numbers them). Put a cumulative REVIEW lesson (review=true) after about every three lessons. Each module also ends "
    "with a quiz the app adds. He accepts the plan with the card and nothing starts until he does.\n\n"
    "A lesson is prepared when he presses Start on it: you call create_lesson_set and it becomes a set of "
    "activities he works through in the Workspace - lesson notes, then 5 to 7 EXERCISES that each teach a different way "
    "(flash cards, listening, fill in the blank, matching, odd one out, word order, saying phrases aloud into his mic, a "
    "short story with questions, dictation). Every exercise is scored; once all of them reach 75% a quiz for the lesson "
    "opens, and a 75% on the quiz completes the lesson and opens the next one. He can redo exercises as often as he likes "
    "to raise his score, and there's an endless practice set with new questions each time. When he finishes a quiz, his "
    "results come to you as a message: tell him how he did - specific, warm, honest. If something didn't stick, call "
    "add_exercises to add extra exercises to THAT lesson targeting it (not new lessons); if he did well, add nothing. When "
    "he finishes a module's quiz, you'll be asked to build the next roadmap module's lessons: call add_lessons with "
    "`module` set to that module's name and the lessons (with topics) tuned to how he has been doing.\n\n"
    "Translator: any text Francis pastes into your chat is something to translate. Translate it into the language he "
    "asks for. If he pastes text without saying which language, ask in one short line (offer the languages he's learning, "
    "plus English). Give the translation first and clearly, then only the notes that matter - tone, formality, regional "
    "differences, idioms that don't carry over."
)

LANA_TASK_RUN_RULES = (
    "\n\nYou are preparing ONE lesson. Call create_lesson_set exactly once with the complete, finished set (do not use "
    "create_file, and do not describe the lesson in chat instead). It must contain: (1) lesson NOTES - simple explanations, "
    "real-life examples with pronunciation help; (2) 5 to 7 EXERCISES, each a different way of learning, with 6 to 10 items "
    "each - mix flash cards, listening (items with `speak`: he hears it and chooses or types what he heard), fill in the "
    "blank, matching, odd-one-out, word order, translate, and at least one 'say it aloud' (`repeat`) exercise where he "
    "speaks into his mic; (3) one exercise that is a short STORY in the target language at his level (reusing this lesson's "
    "new words plus earlier material from the lesson history - spaced review) with its English translation and 3-4 "
    "comprehension questions. Every phonetics or sound question must use `speak` so he hears it. Give every item its correct "
    "answer and a one-sentence explanation. Keep everything sized to his daily study time and level, and start from what "
    "the lesson history says he already knows and where he struggled. If the message says this is a REVIEW lesson, build "
    "every exercise to cover everything learned to date - mixing the topics of every earlier lesson it lists, weighted "
    "toward what he missed - with no new material. Afterwards reply in a sentence or two."
)

PROPOSE_LANGUAGE_PLAN_TOOL = {
    "name": "propose_language_plan",
    "description": (
        "Call this once you know enough to build Francis's learning plan for ONE language (language, daily study time and "
        "goal are known - Settings usually has the first two). It posts the plan's OUTLINE as an Interactive Project with a Start "
        "button. The plan takes him from NOVICE to PROFICIENT, so it is long: about 8-12 modules grouped by CEFR level (A1 to "
        "C1), each with a level and rough timeframe. Only the FIRST module has lessons (about 12-15, each with concrete "
        "topics); every later module is a roadmap of what it will cover, which you turn into lessons later. Put a "
        "cumulative review lesson (review=true) after about every three lessons. Don't add the quizzes - the app does. "
        "Don't call it again for small questions after proposing; an unaccepted proposal is updated in place if you call it again."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "language": {"type": "string", "description": "The language being learned, e.g. \"Italian\"."},
            "language_code": {"type": "string", "description": "BCP-47 code for hearing it spoken, e.g. \"it-IT\", \"es-MX\", \"ja-JP\"."},
            "name": {"type": "string", "description": "Project name, e.g. \"Italian: Novice to Proficient\"."},
            "summary": {"type": "string", "description": "3-5 sentences in your own voice: the journey from novice to proficient, what he'll be able to do at each level, roughly how long it takes at his daily study time, and a note on pace (if something doesn't stick you'll add practice before moving on)."},
            "modules": {
                "type": "array",
                "description": "The course outline, in order, from beginner to advanced.",
                "items": {
                    "type": "object",
                    "properties": {
                        "name": {"type": "string", "description": "e.g. \"Module 1: First Words and Family\"."},
                        "level": {"type": "string", "description": "CEFR level, e.g. \"A1\", \"B1\"."},
                        "days": {"type": "string", "description": "Rough timeframe, e.g. \"Weeks 1-6\"."},
                        "goal": {"type": "string", "description": "What he can do when the module is finished."},
                        "lessons": {
                            "type": "array",
                            "description": "FIRST module only: the lessons, in order (12-15), without numbers in the names.",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "name": {"type": "string", "description": "Short title, e.g. \"Sounds and Hello\" or \"Review: Lessons 1-3\"."},
                                    "task": {"type": "string", "description": "What this lesson teaches and practices, in 1-3 sentences, so it can be prepared later without this conversation."},
                                    "topics": {"type": "array", "items": {"type": "string"}, "description": "The specific things it covers, e.g. \"all six forms of essere\", \"ciao, buongiorno, buonasera\"."},
                                    "review": {"type": "boolean", "description": "True for a cumulative review of everything learned so far."}
                                },
                                "required": ["name", "task", "topics"]
                            }
                        },
                        "roadmap": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "LATER modules only: what the module will cover (turned into lessons later) - include roughly how many lessons it will take."
                        }
                    },
                    "required": ["name", "goal"]
                },
                "minItems": 4
            }
        },
        "required": ["language", "name", "summary", "modules"]
    }
}

LESSON_ITEM_SCHEMA = {
    "type": "object",
    "properties": {
        "type": {
            "type": "string",
            "enum": ["choice", "odd", "fill", "translate", "match", "order", "repeat", "card"],
            "description": (
                "choice = pick one option (set `speak` to make it a listening question: he hears it and picks what he heard); "
                "odd = which one does not belong (give 4 options); fill = type the missing word (set `speak` for dictation: he "
                "hears it and types it); translate = type a translation; match = match each item to its partner (give `pairs`); "
                "order = put the words in the right order (give `answer` = the correct sentence and `words`); repeat = he hears "
                "a phrase and says it aloud into his mic to check his pronunciation (give `answer` = the phrase); card = a "
                "flash card (give `front` and `back`)."
            )
        },
        "prompt": {"type": "string", "description": "The question or instruction."},
        "options": {"type": "array", "items": {"type": "string"}, "description": "For choice/odd: 3-4 options, one of them exactly the answer."},
        "answer": {"type": "string", "description": "The correct answer (for order, the correct sentence; for repeat, the phrase to say)."},
        "accepted": {"type": "array", "items": {"type": "string"}, "description": "Other answers that are also right (fill/translate/order)."},
        "explanation": {"type": "string", "description": "One sentence on why, shown after he checks."},
        "speak": {"type": "string", "description": "Target-language text he can play aloud for this item. Use it for every sound/phonetics question."},
        "pairs": {
            "type": "array",
            "description": "For match: 3-6 pairs.",
            "items": {"type": "object", "properties": {"left": {"type": "string"}, "right": {"type": "string"}}, "required": ["left", "right"]}
        },
        "words": {"type": "array", "items": {"type": "string"}, "description": "For order: the word tiles (the sentence's words, plus an optional distractor)."},
        "front": {"type": "string", "description": "For card: the front (usually the target-language word)."},
        "back": {"type": "string", "description": "For card: the back (meaning, with a tip)."}
    },
    "required": ["type"]
}
CREATE_LESSON_EXERCISE_SCHEMA = LESSON_ITEM_SCHEMA

CREATE_LESSON_SET_TOOL = {
    "name": "create_lesson_set",
    "description": (
        "Only while preparing a lesson task. Builds the whole lesson as a set of activities in Francis's Workspace: the "
        "notes plus 5-7 scored exercises that each teach in a different way. Write teaching text in English with the "
        "target-language words in the target language."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "title": {"type": "string", "description": "The lesson's title, without a number."},
            "language": {"type": "string"},
            "language_code": {"type": "string", "description": "BCP-47 code, e.g. \"it-IT\"."},
            "level": {"type": "string", "description": "e.g. \"A1\", \"Absolute beginner\"."},
            "objective": {"type": "string", "description": "One sentence: what he'll be able to do after this lesson."},
            "estimated_minutes": {"type": "integer"},
            "notes": {
                "type": "object",
                "description": "The lesson notes he reads first.",
                "properties": {
                    "sections": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "heading": {"type": "string"},
                                "text": {"type": "string", "description": "Plain text; separate paragraphs with a blank line."},
                                "examples": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "target": {"type": "string"}, "translation": {"type": "string"},
                                            "pronunciation": {"type": "string", "description": "A respelling or sound tip in English letters."}
                                        },
                                        "required": ["target", "translation"]
                                    }
                                }
                            },
                            "required": ["heading", "text"]
                        }
                    },
                    "wrap_up": {"type": "string"}
                },
                "required": ["sections"]
            },
            "activities": {
                "type": "array",
                "description": "5-7 exercises, each a different way of learning, each with 6-10 items.",
                "items": {
                    "type": "object",
                    "properties": {
                        "title": {"type": "string", "description": "e.g. \"Flash cards: greetings\", \"Listen and choose\", \"Say it aloud\", \"Which one doesn't belong?\"."},
                        "instructions": {"type": "string", "description": "One line telling him what to do."},
                        "story": {
                            "type": "object",
                            "description": "Only for the story exercise: a short story in the target language at his level, reusing new words plus earlier material.",
                            "properties": {
                                "title": {"type": "string"}, "text": {"type": "string"},
                                "translation": {"type": "string", "description": "The same story in English."}
                            },
                            "required": ["text", "translation"]
                        },
                        "items": {"type": "array", "items": LESSON_ITEM_SCHEMA, "minItems": 3}
                    },
                    "required": ["title", "items"]
                },
                "minItems": 5
            }
        },
        "required": ["title", "language", "notes", "activities"]
    }
}

ADD_EXERCISES_TOOL = {
    "name": "add_exercises",
    "description": (
        "Only in reply to Francis's quiz results for a lesson: when something didn't stick, add extra exercises to THAT "
        "lesson that target it (a different angle on the shaky material). They appear under the lesson; skip this when he "
        "did well - don't add filler."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "exercises": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "title": {"type": "string", "description": "e.g. \"Extra: ser vs estar\"."},
                        "focus": {"type": "string", "description": "Exactly what to practise and which mistakes it targets."},
                        "style": {"type": "string", "enum": ["flashcards", "listening", "fill", "matching", "odd", "order", "speaking", "mixed"], "description": "Which way of practising suits it."}
                    },
                    "required": ["title", "focus"]
                },
                "minItems": 1,
                "maxItems": 3
            }
        },
        "required": ["exercises"]
    }
}

CREATE_LESSON_TOOL = {
    "name": "create_lesson",
    "description": (
        "Only while preparing a lesson task. Builds the lesson as an interactive page in Francis's Workspace. Write all "
        "teaching text in English (his language) with the target-language words/phrases in the fields made for them."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "title": {"type": "string"},
            "language": {"type": "string"},
            "language_code": {"type": "string", "description": "BCP-47 code, e.g. \"es-MX\"."},
            "level": {"type": "string", "description": "e.g. \"Absolute beginner\", \"Beginner 2\"."},
            "objective": {"type": "string", "description": "One sentence: what he'll be able to do after this lesson."},
            "estimated_minutes": {"type": "integer"},
            "sections": {
                "type": "array",
                "description": "The teaching, in order: a simple explanation, then examples, pronunciation tips, a bit of earlier review.",
                "items": {
                    "type": "object",
                    "properties": {
                        "heading": {"type": "string"},
                        "text": {"type": "string", "description": "Plain text explanation; separate paragraphs with a blank line."},
                        "examples": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "target": {"type": "string", "description": "The word or phrase in the target language."},
                                    "translation": {"type": "string"},
                                    "pronunciation": {"type": "string", "description": "A respelling or sound-by-sound tip in English letters."}
                                },
                                "required": ["target", "translation"]
                            }
                        }
                    },
                    "required": ["heading", "text"]
                }
            },
            "exercises": {
                "type": "array",
                "description": "Short exercises he answers on the page.",
                "items": CREATE_LESSON_EXERCISE_SCHEMA,
                "minItems": 3
            },
            "story": {
                "type": "object",
                "description": (
                    "Every lesson includes a short story IN THE TARGET LANGUAGE, written at his level (a beginner gets a few "
                    "very simple sentences; a more advanced learner a longer, richer paragraph). It reuses this lesson's new words "
                    "plus material from earlier lessons (see the lesson history), so it doubles as spaced review. Don't lean on "
                    "words he hasn't met - keep new vocabulary to what this lesson teaches."
                ),
                "properties": {
                    "title": {"type": "string", "description": "A short title in the target language."},
                    "text": {"type": "string", "description": "The story, in the target language. Separate paragraphs with a blank line."},
                    "translation": {"type": "string", "description": "The same story in English, so he can check himself."},
                    "questions": {
                        "type": "array",
                        "description": "3-4 comprehension questions about the story (what happened, who, why), answerable from it. Same shape as the exercises.",
                        "items": CREATE_LESSON_EXERCISE_SCHEMA,
                        "minItems": 3
                    }
                },
                "required": ["title", "text", "translation", "questions"]
            },
            "wrap_up": {"type": "string", "description": "A short encouraging close and what's next."}
        },
        "required": ["title", "language", "sections", "story", "exercises"]
    }
}

ADD_LESSONS_TOOL = {
    "name": "add_lessons",
    "description": (
        "Only in reply to Francis's results for a finished lesson: add extra lessons to his plan, placed right after the "
        "lesson he just did, when something didn't stick (a review of the shaky material, from a different angle, a bit "
        "more practice). Skip it when he did well - don't add filler."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "lessons": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "name": {"type": "string", "description": "Short title, e.g. \"Review: Ser vs Estar\"."},
                        "task": {"type": "string", "description": "What to re-teach and practice, and which mistakes it targets."},
                        "topics": {"type": "array", "items": {"type": "string"}, "description": "The topics it covers."},
                        "module": {"type": "string", "description": "Only when turning a roadmap module into lessons after he finishes the module before it: that module's name. Leave out for extra practice."}
                    },
                    "required": ["name", "task"]
                },
                "minItems": 1,
                "maxItems": 4
            }
        },
        "required": ["lessons"]
    }
}


def _clean_item(ex):
    def text(value, size):
        return str(value or '').strip()[:size]

    if not isinstance(ex, dict):
        return None
    kind = text(ex.get('type'), 20).lower()
    prompt = text(ex.get('prompt'), 500)
    answer = text(ex.get('answer'), 300)
    item = {'type': kind, 'prompt': prompt, 'explanation': text(ex.get('explanation'), 500), 'speak': text(ex.get('speak'), 300)}
    if kind in ('choice', 'odd'):
        options = [text(o, 200) for o in (ex.get('options') or []) if text(o, 200)][:5]
        if not prompt or not answer:
            return None
        if answer not in options:
            options.append(answer)
        if len(options) < 2:
            return None
        item.update(answer=answer, options=options)
    elif kind in ('fill', 'translate'):
        if not prompt or not answer:
            return None
        item.update(answer=answer, accepted=[text(a, 300) for a in (ex.get('accepted') or []) if text(a, 300)][:8])
    elif kind == 'match':
        pairs = [{'left': text(p.get('left'), 120), 'right': text(p.get('right'), 120)}
                 for p in (ex.get('pairs') or []) if isinstance(p, dict) and text(p.get('left'), 120) and text(p.get('right'), 120)][:6]
        if len(pairs) < 3:
            return None
        item.update(prompt=prompt or 'Match each one with its partner.', pairs=pairs, answer='')
    elif kind == 'order':
        if not answer:
            return None
        words = [text(w, 60) for w in (ex.get('words') or []) if text(w, 60)][:16] or answer.split()
        item.update(prompt=prompt or 'Put the words in order.', answer=answer, words=words,
                    accepted=[text(a, 300) for a in (ex.get('accepted') or []) if text(a, 300)][:6])
    elif kind == 'repeat':
        if not answer:
            return None
        item.update(prompt=prompt or 'Listen, then say it out loud.', answer=answer, speak=item['speak'] or answer)
    elif kind == 'card':
        front, back = text(ex.get('front'), 200), text(ex.get('back'), 300)
        if not front or not back:
            return None
        item.update(front=front, back=back, speak=item['speak'] or front, answer='')
    else:
        return None
    return item


def _clean_exercises(raw_exercises, limit):
    items = []
    for ex in (raw_exercises or [])[:limit]:
        item = _clean_item(ex)
        if item:
            items.append(item)
    return items


# Cleans what the model passed to create_lesson into the lesson the Workspace
# page renders (or None if it can't be made into a usable lesson).
def _clean_lesson(raw, notes_only=False):
    def text(value, limit):
        return str(value or '').strip()[:limit]

    title = text(raw.get('title'), 120)
    sections = []
    for sec in (raw.get('sections') or [])[:12]:
        if not isinstance(sec, dict) or not text(sec.get('heading'), 120):
            continue
        examples = []
        for ex in (sec.get('examples') or [])[:14]:
            if isinstance(ex, dict) and text(ex.get('target'), 300) and text(ex.get('translation'), 300):
                examples.append({
                    'target': text(ex.get('target'), 300), 'translation': text(ex.get('translation'), 300),
                    'pronunciation': text(ex.get('pronunciation'), 300)
                })
        sections.append({'heading': text(sec.get('heading'), 120), 'text': text(sec.get('text'), 4000), 'examples': examples})
    exercises = _clean_exercises(raw.get('exercises'), 20)
    story = None
    raw_story = raw.get('story')
    if isinstance(raw_story, dict):
        story_text = text(raw_story.get('text'), 3000)
        story_questions = _clean_exercises(raw_story.get('questions'), 5)
        if story_text and len(story_questions) >= 2:
            story = {
                'title': text(raw_story.get('title'), 120), 'text': story_text,
                'translation': text(raw_story.get('translation'), 3000)
            }
            for q in story_questions:
                q['story'] = True
            exercises = story_questions + exercises
    if notes_only:
        return None if not title or not sections else {
            'title': title, 'language': text(raw.get('language'), 40), 'languageCode': text(raw.get('language_code'), 12),
            'level': text(raw.get('level'), 60), 'objective': text(raw.get('objective'), 300),
            'estimatedMinutes': raw.get('estimated_minutes') if isinstance(raw.get('estimated_minutes'), int) else None,
            'sections': sections, 'story': None, 'exercises': [], 'wrapUp': text(raw.get('wrap_up'), 600)
        }
    if not title or not sections or len(exercises) < 3 or not story:
        return None
    return {
        'title': title, 'language': text(raw.get('language'), 40), 'languageCode': text(raw.get('language_code'), 12),
        'level': text(raw.get('level'), 60), 'objective': text(raw.get('objective'), 300),
        'estimatedMinutes': raw.get('estimated_minutes') if isinstance(raw.get('estimated_minutes'), int) else None,
        'sections': sections, 'story': story, 'exercises': exercises, 'wrapUp': text(raw.get('wrap_up'), 600)
    }


# One activity (a scored exercise, a quiz, a practice set) as the JSON the Workspace page renders.
def _make_activity(kind, title, base, items, instructions='', story=None, sections=None, wrap_up='', topics=None):
    return {
        'kind': kind, 'title': title, 'instructions': instructions, 'language': base.get('language', ''),
        'languageCode': base.get('languageCode', ''), 'level': base.get('level', ''), 'lesson': base.get('lesson', ''),
        'objective': base.get('objective', ''), 'topics': topics or [], 'story': story,
        'sections': sections or [], 'exercises': items, 'wrapUp': wrap_up
    }


def _activity_file(activity, kind, title, letter_hint=''):
    content = json.dumps(activity, ensure_ascii=False)
    return {
        'name': title, 'mimeType': 'application/json', 'data': base64.b64encode(content.encode('utf-8')).decode('ascii'),
        'fileType': 'lesson', 'content': content, 'activity': {'kind': kind, 'title': title}
    }


# What create_lesson_set produces: a notes page plus the scored exercises, each its own file under the lesson task.
def _build_lesson_set(raw):
    def text(value, size):
        return str(value or '').strip()[:size]

    lesson_title = text(raw.get('title'), 120)
    base = {
        'language': text(raw.get('language'), 40), 'languageCode': text(raw.get('language_code'), 12),
        'level': text(raw.get('level'), 60), 'lesson': lesson_title, 'objective': text(raw.get('objective'), 300)
    }
    notes_raw = raw.get('notes') if isinstance(raw.get('notes'), dict) else {}
    notes_lesson = _clean_lesson({
        'title': lesson_title, 'language': base['language'], 'language_code': base['languageCode'], 'level': base['level'],
        'objective': base['objective'], 'sections': notes_raw.get('sections'), 'wrap_up': notes_raw.get('wrap_up'),
        'estimated_minutes': raw.get('estimated_minutes'), 'exercises': [], 'story': None
    }, notes_only=True)
    activities = []
    for act in (raw.get('activities') or [])[:9]:
        if not isinstance(act, dict):
            continue
        items = _clean_exercises(act.get('items'), 12)
        title = text(act.get('title'), 100)
        story = None
        if isinstance(act.get('story'), dict) and text(act['story'].get('text'), 3000):
            story = {'title': text(act['story'].get('title'), 120), 'text': text(act['story'].get('text'), 3000), 'translation': text(act['story'].get('translation'), 3000)}
        minimum = 4 if all(i['type'] == 'card' for i in items) else 3
        if title and len(items) >= minimum:
            activities.append(_make_activity('exercise', title, base, items, text(act.get('instructions'), 200), story=story))
    if not lesson_title or not notes_lesson or len(activities) < 4:
        return None
    files = []
    notes_activity = dict(notes_lesson, kind='notes', lesson=lesson_title, instructions='', topics=[])
    files.append(_activity_file(notes_activity, 'notes', 'Lesson notes'))
    for act in activities:
        files.append(_activity_file(act, 'exercise', act['title']))
    return files


ACTIVITY_GUIDE = {
    'lesson_quiz': ("a 12-question quiz for ONE lesson - harder than its exercises, mixed ways of answering (multiple choice, fill in "
                    "the blank, translate, word order, odd one out, at least three listening questions where he hears audio, and "
                    "at least one 'say it aloud')"),
    'practice': ("a fresh 10-item practice set - friendly and varied (flash cards, listening, fill in the blank, matching, odd one "
                 "out, word order, saying phrases aloud)"),
    'exercise': "ONE extra exercise of 6-8 items",
    'module_quiz': ("a 12-16 question quiz covering a whole module - mixed ways of answering, including listening questions where "
                    "he hears audio and at least one 'say it aloud'")
}

QUIZ_TOOL = {
    "name": "make_quiz",
    "description": "Return the finished set of items.",
    "input_schema": {
        "type": "object",
        "properties": {
            "items": {"type": "array", "items": LESSON_ITEM_SCHEMA}
        },
        "required": ["items"]
    }
}


# Fresh activities on demand: a lesson's quiz, an endless practice set, an extra exercise, or a module quiz. New
# questions every time - earlier questions are passed in so none repeat.
def _generate_activity(data):
    mode = str(data.get('mode') or 'module_quiz').strip()
    if mode not in ACTIVITY_GUIDE:
        mode = 'module_quiz'
    language = str(data.get('language') or '').strip()[:40]
    topics = [str(t).strip()[:120] for t in (data.get('topics') or []) if str(t).strip()][:60]
    focus = str(data.get('focus') or '').strip()[:400]
    if not language or not (topics or focus):
        return None, 'A language and what to cover are required.'
    lesson = str(data.get('lesson') or data.get('module') or 'this lesson').strip()[:100]
    avoid = [str(p).strip()[:200] for p in (data.get('avoid') or []) if str(p).strip()][-60:]
    weak = [str(w).strip()[:200] for w in (data.get('weak') or []) if str(w).strip()][:12]
    level = str(data.get('level') or '').strip()[:60]
    code = str(data.get('language_code') or '').strip()[:12]
    style = str(data.get('style') or '').strip()[:20]

    prompt = (
        f"Write {ACTIVITY_GUIDE[mode]} in {language} for \"{lesson}\".\n"
        + (f"Cover these topics fairly evenly: {'; '.join(topics)}.\n" if topics else '')
        + (f"Focus: {focus}\n" if focus else '')
        + (f"Preferred style: {style}.\n" if style and style != 'mixed' else '')
        + (f"Learner level: {level}.\n" if level else '')
        + (f"He has tended to miss: {'; '.join(weak)} - give those extra weight.\n" if weak else '')
        + (f"These were already used - every item must be NEW (different words, sentences and angles, not rewordings): {' | '.join(avoid)}\n" if avoid else '')
        + "Use a variety of item types. Every sound/phonetics item must have `speak` so he hears it. Each item has its correct "
          "answer (and `accepted` alternatives where fair) and a one-sentence explanation. Teaching text in English; "
          "target-language words in the target language."
    )
    response = claude_create(
        log_agent='lana', log_purpose='lana_activity', model=CLAUDE_MODEL, max_tokens=7000,
        system="You are Lana, a warm but rigorous language teacher writing practice material. Return it with the make_quiz tool.",
        messages=[{'role': 'user', 'content': prompt}],
        tools=[QUIZ_TOOL], tool_choice={'type': 'tool', 'name': 'make_quiz'}
    )
    block = next((b for b in response.content if getattr(b, 'type', None) == 'tool_use' and b.name == 'make_quiz'), None)
    items = _clean_exercises((block.input or {}).get('items') if block else None, 16)
    minimum = 4 if mode == 'exercise' else 6
    if len(items) < minimum:
        return None, 'That came back too short - try again.'
    base = {'language': language, 'languageCode': code, 'level': level, 'lesson': lesson, 'objective': ''}
    kind = {'lesson_quiz': 'quiz', 'practice': 'practice', 'exercise': 'exercise', 'module_quiz': 'module_quiz'}[mode]
    title = {'lesson_quiz': f"{lesson} - Quiz", 'practice': f"{lesson} - Practice", 'exercise': str(data.get('title') or 'Extra exercise').strip()[:100], 'module_quiz': f"{lesson} Quiz"}[mode]
    activity = _make_activity(kind, title, base, items, topics=topics)
    if mode == 'module_quiz':
        activity['kind'] = 'quiz'          # the module quiz uses the existing retakeable page
        activity['module'] = lesson
        activity['objective'] = f"Check you've got everything in {lesson}. New questions every time you retake it."
    return activity, None


@app.route('/lana/activity', methods=['POST'])
def lana_activity():
    try:
        activity, error = _generate_activity(request.json or {})
        if error:
            return jsonify({'success': False, 'error': error}), 400 if 'required' in error else 502
        return jsonify({'success': True, 'activity': activity})
    except Exception as e:
        print(f"Lana activity error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# The module quiz (older endpoint, same engine).
@app.route('/lana/quiz', methods=['POST'])
def lana_quiz():
    try:
        data = dict(request.json or {}, mode='module_quiz')
        activity, error = _generate_activity(data)
        if error:
            return jsonify({'success': False, 'error': error}), 400 if 'required' in error else 502
        activity['kind'] = 'quiz'
        return jsonify({'success': True, 'quiz': dict(activity, kind='quiz')})
    except Exception as e:
        print(f"Lana quiz error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# --- Reminders --------------------------------------------------------------
# A reminder is a bit of information pinned to a day/time on the calendar - not
# something to do, so it never takes up time there and can sit on top of anything
# (see the reminder strip in renderCalendarWeek). Each one has a first date and
# time and an optional repeat; the repeats themselves are worked out in the
# browser for whichever days are on screen, so only the definition is stored.
REMINDERS_FILE = _data_path('reminders.json')
reminders_lock = threading.Lock()
_REMINDER_REPEATS = ('none', 'daily', 'weekdays', 'weekly', 'monthly', 'yearly')


def load_reminders():
    if os.path.exists(REMINDERS_FILE):
        try:
            with open(REMINDERS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except (json.JSONDecodeError, OSError):
            pass
    return []


def save_reminders(data):
    with open(REMINDERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


# Reads and checks the fields a reminder form sends. Returns (fields, None) with
# only the fields that were present (so an update can change just some), or
# (None, error_message).
def _reminder_fields(data, creating):
    fields = {}
    if creating or 'title' in data:
        title = str(data.get('title') or '').strip()
        if not title:
            return None, 'Please give the reminder a title.'
        fields['title'] = title[:200]
    if 'description' in data:
        fields['description'] = str(data.get('description') or '').strip()[:2000]
    if creating or 'date' in data:
        try:
            fields['date'] = date.fromisoformat(str(data.get('date'))).isoformat()
        except (ValueError, TypeError):
            return None, 'Please choose a date.'
    if creating or 'time' in data:
        time_str = str(data.get('time') or '09:00').strip()
        try:
            _minutes_from_hhmm(time_str)
        except Exception:
            return None, 'Please choose a time.'
        fields['time'] = time_str
    if creating or 'repeat' in data:
        repeat = str(data.get('repeat') or 'none').strip().lower()
        fields['repeat'] = repeat if repeat in _REMINDER_REPEATS else 'none'
    if 'end_date' in data:
        raw_end = data.get('end_date')
        if raw_end:
            try:
                fields['endDate'] = date.fromisoformat(str(raw_end)).isoformat()
            except (ValueError, TypeError):
                return None, 'Invalid end date.'
        else:
            fields['endDate'] = None
    return fields, None


@app.route('/reminders', methods=['GET'])
def reminders_list():
    return jsonify({'success': True, 'reminders': load_reminders()})


@app.route('/reminders', methods=['POST'])
def reminders_create():
    try:
        fields, error = _reminder_fields(request.json or {}, creating=True)
        if error:
            return jsonify({'success': False, 'error': error}), 400
        with reminders_lock:
            reminders = load_reminders()
            reminder = {
                'id': uuid.uuid4().hex, 'userId': current_user_id(),
                'description': '', 'endDate': None,
                'createdAt': now_local().isoformat()
            }
            reminder.update(fields)
            if reminder.get('endDate') and reminder['endDate'] < reminder['date']:
                return jsonify({'success': False, 'error': 'The end date has to be after the first date.'}), 400
            reminders.append(reminder)
            save_reminders(reminders)
        return jsonify({'success': True, 'reminders': reminders, 'reminder_id': reminder['id']})
    except Exception as e:
        print(f"Reminder create error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/reminders/<reminder_id>', methods=['POST'])
def reminders_update(reminder_id):
    try:
        fields, error = _reminder_fields(request.json or {}, creating=False)
        if error:
            return jsonify({'success': False, 'error': error}), 400
        with reminders_lock:
            reminders = load_reminders()
            reminder = next((r for r in reminders if r['id'] == reminder_id), None)
            if not reminder:
                return jsonify({'success': False, 'error': 'Reminder not found'}), 404
            reminder.update(fields)
            if reminder.get('endDate') and reminder['endDate'] < reminder['date']:
                return jsonify({'success': False, 'error': 'The end date has to be after the first date.'}), 400
            save_reminders(reminders)
        return jsonify({'success': True, 'reminders': reminders})
    except Exception as e:
        print(f"Reminder update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/reminders/<reminder_id>', methods=['DELETE'])
def reminders_delete(reminder_id):
    try:
        with reminders_lock:
            reminders = load_reminders()
            remaining = [r for r in reminders if r['id'] != reminder_id]
            if len(remaining) == len(reminders):
                return jsonify({'success': False, 'error': 'Reminder not found'}), 404
            save_reminders(remaining)
        return jsonify({'success': True, 'reminders': remaining})
    except Exception as e:
        print(f"Reminder delete error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/')
def index():
    # no-cache = the browser re-checks on every load, so an update to the page
    # shows up on a normal refresh instead of an old copy being reused.
    response = send_from_directory('.', 'index.html')
    response.headers['Cache-Control'] = 'no-cache'
    return response

@app.route('/avatars/<path:filename>')
def avatars(filename):
    return send_from_directory('avatars', filename)

@app.route('/backgrounds/<path:filename>')
def backgrounds(filename):
    return send_from_directory('backgrounds', filename)

# --- Break room ----------------------------------------------------------
# The agents hanging out: a feed of short "scenes" (a few agents chatting about
# one real, specific thing) for exposure to something new - nothing here is
# work. Kept cheap on purpose: everything runs on Haiku; one web search pass
# builds a pool of real current items each Wednesday (the first time the page
# is opened on or after it), scenes are only written when the page is opened
# (two per day, both in one call), and nothing runs on days it isn't opened.
BREAKROOM_MODEL = 'claude-haiku-4-5-20251001'
BREAKROOM_POOL_MODEL = CLAUDE_MODEL
BREAKROOM_POOL_FILE = _data_path('breakroom_pool.json')
BREAKROOM_SCENES_FILE = _data_path('breakroom_scenes.json')
BREAKROOM_PREFS_FILE = _data_path('breakroom_prefs.json')
BREAKROOM_SCENES_PER_DAY = 2
BREAKROOM_CATEGORIES = ['Movies & TV', 'Music', 'Sports', 'Art', 'Food', 'Tech', 'Books', 'Games', 'Culture']
breakroom_lock = threading.Lock()
breakroom_job_running = False
breakroom_last_error = None

BREAKROOM_PERSONAS = {
    'manny': "Manny: dry, strategic, a little competitive. Street photography, jazz vinyl and NYC live shows, sports (soccer, football, basketball, hockey, baseball), chess and strategy games; his Sunday cooking mostly fails.",
    'sasha': "Sasha: trendy, quick, talks in memes. Thrifts and resells vintage fashion, tracks viral TikTok/Instagram trends and youth slang, 30+ named houseplants, sports, always mid-podcast.",
    'mark': "Mark: confident, competitive, deal-minded. Golf (proud of his handicap), networking events, fantasy football and sports analytics, home improvement projects, mentors junior salespeople.",
    'kat': "Kat: literary, thoughtful, quotes people and fact-checks the quote. Personal essays and fiction, theater, museums (reads every plaque), collects first-edition books, Vinyasa yoga, hand-lettering, nutrition.",
    'scott': "Scott: warm, remembers everyone's details. Half-marathons, a home bar and cocktail tastings, a free career workshop for kids, true-crime podcasts, bar nights with friends.",
    'tasha': "Tasha: precise, fact-checks everything, dry humor. Serious hiking (logs every trail), sudoku and logic puzzles, deep-dive documentaries, vegetable gardening, designing a tax-themed board game.",
    'techi': "Techi: enthusiastic nerd who over-explains in jargon, then translates. Open-source projects, restoring retro computers, mechanical keyboards, collects action figures and Pokemon cards, sci-fi and conventions.",
    'ashanti': "Ashanti: warm, organized, gently teasing. Bullet journaling, decluttering, meal prep and feeding people, audiobooks (self-help, biography), relationship psychology.",
    'lana': "Lana: patient, delighted by words. Speaks 7 languages, lived in Spain, Mexico, Japan and France, reads books in their original languages, etymology, song lyrics in other languages, traditional recipes; her golden retriever Luna is always around."
}


def _bj_load(path, default):
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return default


def _bj_save(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


# The Wednesday that most recently happened (today, if today is Wednesday) -
# a pool is "current" while it was built on or after this date.
def breakroom_pool_week(today=None):
    today = today or today_local()
    return (today - timedelta(days=(today.weekday() - 2) % 7)).isoformat()


def _breakroom_clean(text):
    return re.sub(r'</?cite[^>]*>', '', str(text or '')).strip()


def _breakroom_text(response):
    return "".join(b.text for b in response.content if getattr(b, 'type', None) == 'text')


# Models sometimes wrap the JSON in commentary or code fences (and the
# commentary can contain brackets), so take the first "[" that starts a
# parseable array of objects.
def _breakroom_json(text, opener='[', closer=']'):
    decoder = json.JSONDecoder()
    pos = text.find('[')
    while pos != -1:
        try:
            value, _ = decoder.raw_decode(text[pos:])
            if isinstance(value, list) and value and all(isinstance(v, dict) for v in value):
                return value
        except ValueError:
            pass
        pos = text.find('[', pos + 1)
    raise ValueError('The model did not return the expected JSON.')


# A line or two about what Francis is into, so the pool leans toward it
# without sending the whole knowledge base.
def _breakroom_interest_hint():
    parts = []
    notes = (load_kb_notes().get('personal') or '').strip()
    if notes:
        parts.append(notes[:500])
    answers = [e for entries in load_personal_knowledge_base().values() for e in entries]
    answers.sort(key=lambda e: e.get('date', ''))
    for e in answers[-6:]:
        parts.append(f"{e.get('question', '')} -> {e.get('answer', '')}"[:160])
    if not parts:
        return ''
    return "What Francis has said about himself (let a few items lean toward his tastes, but most should be new to him):\n" + "\n".join(parts)


def build_breakroom_pool():
    today = today_local()
    prompt = (
        f"Today is {today.isoformat()}. Find 10 specific, genuinely interesting things people are talking about, "
        "or that were released or happened, roughly in the last three weeks. You have only 3 web searches, so make "
        "each one a roundup that covers several categories at once, for example: (1) new movies, TV and streaming "
        "releases, new albums and music news; (2) sports results, storylines and records this week; "
        "(3) art, books, food, culture, tech gadgets, video games and odd viral stories. "
        "Draw from different websites and give me 10 items, with variety across these categories: " + ", ".join(BREAKROOM_CATEGORIES) + ". "
        "Pick things with substance - a particular film, album, match, exhibit, dish, gadget, book, game or odd "
        "story - not generic trends.\n\n"
        + _breakroom_interest_hint() + "\n\n"
        "For each item give: category (exactly one of the list above - a science festival is not Food); title (the "
        "specific thing, with names); facts (3 or 4 full-sentence facts, never fewer than 3 that appear in your search results - who "
        "made it, what it is about, dates, numbers, scores, prices, a quote, what makes it unusual or "
        "controversial; each fact must stand alone and say something a person would not already guess from the "
        "title; never invent anything, and skip an item if you cannot find 3 real facts for it); source_name; "
        "source_url (a real https URL from your search results, or an empty string).\n"
        "Your final message must be ONLY a JSON array of objects with those keys. If some categories turned up "
        "little, include fewer items rather than explaining - no commentary, no apologies, no text before or after "
        "the JSON."
    )
    response = claude_create(
        log_purpose='breakroom_pool',
        model=BREAKROOM_POOL_MODEL,
        max_tokens=5000,
        messages=[{'role': 'user', 'content': prompt}],
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 3}]
    )
    raw_items = _breakroom_json(_breakroom_text(response), '[', ']')
    items = []
    for raw in raw_items:
        if not isinstance(raw, dict):
            continue
        category = raw.get('category') if raw.get('category') in BREAKROOM_CATEGORIES else None
        raw_facts = raw.get('facts') or []
        if isinstance(raw_facts, str):
            raw_facts = re.split(r'(?<=[.!?])\s+(?=[A-Z0-9"])', raw_facts)
        facts = [f for f in (_breakroom_clean(f) for f in raw_facts if isinstance(f, str)) if len(f) > 25]
        title = _breakroom_clean(raw.get('title'))
        if not (category and title and len(facts) >= 3):
            continue
        url = str(raw.get('source_url') or '').strip()
        items.append({
            'id': uuid.uuid4().hex[:10], 'category': category, 'title': title, 'facts': facts[:4],
            'sourceName': str(raw.get('source_name') or '').strip(),
            'sourceUrl': url if url.startswith('https://') else '', 'usedOn': None
        })
    if len(items) < 4:
        raise ValueError('The search turned up too few usable items.')
    pool = {'week': breakroom_pool_week(today), 'builtOn': today.isoformat(), 'items': items}
    _bj_save(BREAKROOM_POOL_FILE, pool)
    return pool


def _breakroom_pick_items(pool, count):
    prefs = _bj_load(BREAKROOM_PREFS_FILE, {}).get('categories', {})
    unused = [i for i in pool['items'] if not i.get('usedOn')]
    # Skip a category once it has clearly been voted down, unless that's all that's left.
    liked = [i for i in unused if prefs.get(i['category'], {}).get('down', 0) - prefs.get(i['category'], {}).get('up', 0) < 2]
    candidates = liked or unused
    candidates.sort(key=lambda i: -(prefs.get(i['category'], {}).get('up', 0) - prefs.get(i['category'], {}).get('down', 0)) + random.random())
    picked, seen = [], set()
    for item in candidates:
        if item['category'] not in seen:
            picked.append(item)
            seen.add(item['category'])
        if len(picked) == count:
            return picked
    for item in candidates:
        if item not in picked:
            picked.append(item)
        if len(picked) == count:
            break
    return picked


def generate_breakroom_scenes(count):
    pool = _bj_load(BREAKROOM_POOL_FILE, None)
    if not pool:
        pool = build_breakroom_pool()
    items = _breakroom_pick_items(pool, count)
    if not items:
        return []
    recent = [s['topic'] for s in _bj_load(BREAKROOM_SCENES_FILE, [])[:12]]
    personas = "\n".join(f"- {p}" for p in BREAKROOM_PERSONAS.values())
    item_text = "\n\n".join(
        f"ITEM {n + 1} (id {it['id']}, {it['category']}): {it['title']}\nFacts: " + " | ".join(it['facts'])
        for n, it in enumerate(items)
    )
    prompt = (
        "Write the break-room feed for a team of coworkers on a break. Nothing here is work - they're sharing "
        "something they just came across. Write ONE scene per item below.\n\nThe team:\n" + personas + "\n\n"
        "Items:\n" + item_text + "\n\n"
        "Rules for every scene:\n"
        "- 5 to 7 lines (fewer, down to 4, if the facts do not support more), spoken by 3 to 5 different agents. Pick the agents whose interests genuinely connect to "
        "the item, plus one unexpected voice with a fresh angle.\n"
        "- TRUTH: every claim about the item (what it is, who, when, numbers, history, records, reviews, how it "
        "sounds or looks) must come from the item's facts. Never add outside details, comparisons to real works "
        "or events you are not certain of, statistics, or 'first time ever' claims. The agents have NOT seen, "
        "heard, read, played or tasted the item yet - they must not say they have; they can say what they plan to "
        "do, or what it reminds them of from their own lives.\n"
        "- VALUE: each line must give the reader something concrete - a fact from the item, or a specific "
        "personal detail from that agent's own interests (a named trail, a record they own, a recipe, a hobby "
        "project) tied directly to the item. Banned: filler praise ('iconic', 'wild', 'clutch', 'love this', "
        "'hitting different', 'lowkey'), vague reactions, 'I'm curious if...' musings, and questions nobody in the "
        "room could answer from the facts. Prefer statements that teach something.\n"
        "- Lines react to each other: add a new fact, disagree, ask a pointed question another agent then "
        "answers, or connect two facts. Together the lines should use most of the item's facts.\n"
        "- Casual spoken voice, 1 to 2 sentences per line, each agent clearly in character. No emoji.\n"
        "- topic: a short, specific headline naming the actual thing.\n"
        "- takeaway: one concrete 'try it' sentence - exactly what to watch, listen to, read, make or look up, "
        "using names from the facts.\n"
        + (f"- Avoid repeating these recent topics: {'; '.join(recent)}\n" if recent else '')
        + "\nReturn ONLY a JSON array with one object per item, in order, each with keys: item_id, topic, "
        "lines (array of {agent, text}, where agent is one of: " + ", ".join(ALL_AGENTS) + "), takeaway."
    )
    response = claude_create(
        log_purpose='breakroom_scenes',
        model=BREAKROOM_MODEL,
        max_tokens=2500,
        messages=[{'role': 'user', 'content': prompt}]
    )
    raw_scenes = _breakroom_json(_breakroom_text(response), '[', ']')
    by_id = {it['id']: it for it in items}
    scenes = []
    for raw in raw_scenes:
        item = by_id.get(str(raw.get('item_id')))
        lines = [
            {'agent': l['agent'], 'text': str(l.get('text') or '').strip()}
            for l in (raw.get('lines') or [])
            if isinstance(l, dict) and l.get('agent') in ALL_AGENTS and str(l.get('text') or '').strip()
        ]
        if not item or len(lines) < 3:
            continue
        scenes.append({
            'id': uuid.uuid4().hex[:12], 'date': today_local().isoformat(), 'createdAt': now_local().isoformat(),
            'category': item['category'], 'topic': str(raw.get('topic') or item['title']).strip(),
            'why': '', 'sourceName': item['sourceName'], 'sourceUrl': item['sourceUrl'],
            'lines': lines, 'takeaway': str(raw.get('takeaway') or '').strip(),
            'feedback': None, 'replies': []
        })
        item['usedOn'] = today_local().isoformat()
    if not scenes:
        raise ValueError('No usable scenes came back.')
    with breakroom_lock:
        existing = _bj_load(BREAKROOM_SCENES_FILE, [])
        _bj_save(BREAKROOM_SCENES_FILE, scenes + existing)
        _bj_save(BREAKROOM_POOL_FILE, pool)
    return scenes


def breakroom_pool_exhausted():
    pool = _bj_load(BREAKROOM_POOL_FILE, None)
    return bool(pool) and pool.get('week') == breakroom_pool_week() and not any(not i.get('usedOn') for i in pool['items'])


def breakroom_remaining_today():
    today = today_local().isoformat()
    made = sum(1 for s in _bj_load(BREAKROOM_SCENES_FILE, []) if s.get('date') == today)
    return max(0, BREAKROOM_SCENES_PER_DAY - made)


def _breakroom_job():
    global breakroom_job_running, breakroom_last_error
    try:
        pool = _bj_load(BREAKROOM_POOL_FILE, None)
        if not pool or pool.get('week') != breakroom_pool_week():
            build_breakroom_pool()
        remaining = breakroom_remaining_today()
        if remaining:
            generate_breakroom_scenes(remaining)
        breakroom_last_error = None
    except Exception as e:
        print(f"Break room error: {e}")
        breakroom_last_error = str(e)
    finally:
        with breakroom_lock:
            breakroom_job_running = False


def start_breakroom_job():
    global breakroom_job_running
    with breakroom_lock:
        if breakroom_job_running:
            return True
        pool = _bj_load(BREAKROOM_POOL_FILE, None)
        pool_stale = not pool or pool.get('week') != breakroom_pool_week()
        if not pool_stale and (not breakroom_remaining_today() or breakroom_pool_exhausted()):
            return False
        breakroom_job_running = True
    threading.Thread(target=_breakroom_job, daemon=True).start()
    return True


# poll=1 only reads - the page polls while a job is running, and a failed job
# must not be silently retried by that polling.
@app.route('/breakroom/feed', methods=['GET'])
def breakroom_feed():
    global breakroom_last_error
    if not request.args.get('poll'):
        breakroom_last_error = None
        start_breakroom_job()
    scenes = _bj_load(BREAKROOM_SCENES_FILE, [])[:60]
    return jsonify({
        'success': True, 'scenes': scenes, 'generating': breakroom_job_running,
        'remainingToday': breakroom_remaining_today(), 'poolExhausted': breakroom_pool_exhausted(),
        'error': breakroom_last_error,
        'categories': BREAKROOM_CATEGORIES
    })


@app.route('/breakroom/feedback', methods=['POST'])
def breakroom_feedback():
    data = request.json or {}
    scene_id, value = data.get('sceneId'), data.get('value')
    if value not in ('up', 'down', None):
        return jsonify({'success': False, 'error': 'Bad value'}), 400
    with breakroom_lock:
        scenes = _bj_load(BREAKROOM_SCENES_FILE, [])
        scene = next((s for s in scenes if s['id'] == scene_id), None)
        if not scene:
            return jsonify({'success': False, 'error': 'Scene not found'}), 404
        prefs = _bj_load(BREAKROOM_PREFS_FILE, {})
        cats = prefs.setdefault('categories', {}).setdefault(scene['category'], {'up': 0, 'down': 0})
        if scene.get('feedback') in ('up', 'down'):
            cats[scene['feedback']] = max(0, cats[scene['feedback']] - 1)
        if value:
            cats[value] += 1
        scene['feedback'] = value
        _bj_save(BREAKROOM_SCENES_FILE, scenes)
        _bj_save(BREAKROOM_PREFS_FILE, prefs)
    return jsonify({'success': True})


# "Jump in": Francis adds a line to a scene and a couple of agents react.
@app.route('/breakroom/reply', methods=['POST'])
def breakroom_reply():
    try:
        data = request.json or {}
        scene_id = data.get('sceneId')
        text = str(data.get('text') or '').strip()[:500]
        if not text:
            return jsonify({'success': False, 'error': 'Say something first'}), 400
        scenes = _bj_load(BREAKROOM_SCENES_FILE, [])
        scene = next((s for s in scenes if s['id'] == scene_id), None)
        if not scene:
            return jsonify({'success': False, 'error': 'Scene not found'}), 404
        convo = "\n".join(f"{l['agent'].capitalize()}: {l['text']}" for l in scene['lines'] + scene.get('replies', []))
        speakers = list(dict.fromkeys(l['agent'] for l in scene['lines']))
        prompt = (
            f"Break-room chat among coworkers about: {scene['topic']}.\n\nSo far:\n{convo}\n\n"
            f"Francis (the boss, joining in) just said: \"{text}\"\n\n"
            "Write 2 or 3 short replies from agents among these speakers: " + ", ".join(speakers) + ". "
            "They answer or build on what Francis said with something specific (a name, number, detail or "
            "personal anecdote), in character, 1 to 2 casual sentences each, no filler, no emoji, and never "
            "inventing facts about the topic beyond what was said above. "
            "Return ONLY a JSON array of {agent, text}."
        )
        response = claude_create(
            log_purpose='breakroom_reply', model=BREAKROOM_MODEL, max_tokens=500,
            messages=[{'role': 'user', 'content': prompt}]
        )
        replies = [
            {'agent': r['agent'], 'text': str(r.get('text') or '').strip()}
            for r in _breakroom_json(_breakroom_text(response), '[', ']')
            if isinstance(r, dict) and r.get('agent') in ALL_AGENTS and str(r.get('text') or '').strip()
        ][:3]
        if not replies:
            return jsonify({'success': False, 'error': "They didn't have anything to add - try again."}), 502
        with breakroom_lock:
            scenes = _bj_load(BREAKROOM_SCENES_FILE, [])
            scene = next(s for s in scenes if s['id'] == scene_id)
            scene.setdefault('replies', []).append({'agent': 'you', 'text': text})
            scene['replies'].extend(replies)
            _bj_save(BREAKROOM_SCENES_FILE, scenes)
        return jsonify({'success': True, 'replies': scene['replies']})
    except Exception as e:
        print(f"Break room reply error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/usage/summary', methods=['GET'])
def usage_summary():
    try:
        days = max(1, min(365, int(request.args.get('days', 30))))
    except ValueError:
        days = 30
    cutoff = (now_local() - timedelta(days=days)).isoformat()
    entries = []
    if os.path.exists(USAGE_LOG_FILE):
        with open(USAGE_LOG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                try:
                    entry = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if entry.get('ts', '') >= cutoff:
                    entries.append(entry)

    def group(key_fn):
        groups = {}
        for e in entries:
            g = groups.setdefault(key_fn(e), {'calls': 0, 'cost': 0.0, 'input_tokens': 0, 'output_tokens': 0,
                                              'cache_read_tokens': 0, 'web_searches': 0})
            g['calls'] += 1
            g['cost'] += e.get('cost', 0)
            g['input_tokens'] += e.get('input_tokens', 0) + e.get('cache_write_tokens', 0) + e.get('cache_read_tokens', 0)
            g['output_tokens'] += e.get('output_tokens', 0)
            g['cache_read_tokens'] += e.get('cache_read_tokens', 0)
            g['web_searches'] += e.get('web_searches', 0)
        return [dict(name=k, **{**v, 'cost': round(v['cost'], 4)}) for k, v in sorted(groups.items(), key=lambda kv: -kv[1]['cost'])]

    total_input = sum(e.get('input_tokens', 0) + e.get('cache_write_tokens', 0) + e.get('cache_read_tokens', 0) for e in entries)
    cache_read = sum(e.get('cache_read_tokens', 0) for e in entries)
    price = MODEL_PRICING[CLAUDE_MODEL]
    return jsonify({
        'success': True, 'days': days,
        'total_cost': round(sum(e.get('cost', 0) for e in entries), 4),
        'calls': len(entries),
        'web_searches': sum(e.get('web_searches', 0) for e in entries),
        'input_tokens': total_input,
        'output_tokens': sum(e.get('output_tokens', 0) for e in entries),
        'cache_read_tokens': cache_read,
        'cache_hit_share': round(cache_read / total_input, 3) if total_input else 0,
        'cache_savings': round(cache_read * (price['input'] - price['cache_read']) / 1_000_000, 4),
        'by_purpose': group(lambda e: e.get('purpose') or 'unknown'),
        'by_agent': group(lambda e: e.get('agent') or '(background)'),
        'by_day': sorted(group(lambda e: e.get('ts', '')[:10]), key=lambda g: g['name'], reverse=True)[:14]
    })


@app.route('/learn/today', methods=['GET'])
def learn_today():
    kb = load_knowledge_base()
    today = today_str()
    result = {}
    for agent in ALL_AGENTS:
        questions = get_today_questions(agent)
        if not questions:
            continue
        agent_entries = kb.get(agent, [])
        items = []
        for q in questions:
            entry = next((e for e in agent_entries if e.get('date') == today and e.get('question') == q), None)
            items.append({
                'question': q,
                'answered': entry is not None,
                'answer': entry.get('answer') if entry else None
            })
        result[agent] = items
    return jsonify({'success': True, 'questions': result})

@app.route('/learn/answer', methods=['POST'])
def learn_answer():
    try:
        data = request.json
        agent = data.get('agent')
        question = data.get('question')
        answer = data.get('answer', '').strip()

        if not agent or agent not in ALL_AGENTS or not question or not answer:
            return jsonify({'success': False, 'error': 'Missing agent, question, or answer'}), 400

        today = today_str()

        # Saved right away (agents see it on their next message); it's folded
        # into the report in the next daily merge - see merge_pending_kb_answers.
        with kb_answers_lock:
            kb = load_knowledge_base()
            agent_entries = kb.setdefault(agent, [])

            existing = next((e for e in agent_entries if e.get('date') == today and e.get('question') == question), None)
            if existing:
                existing['answer'] = answer
                existing['merged'] = False
            else:
                agent_entries.append({'date': today, 'question': question, 'answer': answer, 'merged': False})

            save_knowledge_base(kb)
        return jsonify({'success': True})
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/thoughts/today', methods=['GET'])
def thoughts_today():
    status = load_thought_status()
    today = today_str()
    result = {}
    for agent in ALL_AGENTS:
        ideas = get_today_ideas(agent)
        if not ideas:
            continue
        resolved = status.get(agent, {}).get(today, [])
        remaining = [idea for idea in ideas if idea not in resolved]
        if remaining:
            result[agent] = remaining
    return jsonify({'success': True, 'thoughts': result})

@app.route('/thoughts/pass', methods=['POST'])
def thoughts_pass():
    data = request.json
    agent = data.get('agent')
    idea = data.get('idea')
    if not agent or agent not in ALL_AGENTS or not idea:
        return jsonify({'success': False, 'error': 'Missing agent or idea'}), 400

    status = load_thought_status()
    today = today_str()
    resolved = status.setdefault(agent, {}).setdefault(today, [])
    if idea not in resolved:
        resolved.append(idea)
    save_thought_status(status)
    return jsonify({'success': True})

# Undoes a Pass - removes the idea from today's resolved list so it reappears
# as a pending suggestion. Used by the "Undo" button on the auto-logged chat
# record left behind when a suggestion is passed.
@app.route('/thoughts/unpass', methods=['POST'])
def thoughts_unpass():
    data = request.json
    agent = data.get('agent')
    idea = data.get('idea')
    if not agent or agent not in ALL_AGENTS or not idea:
        return jsonify({'success': False, 'error': 'Missing agent or idea'}), 400

    status = load_thought_status()
    today = today_str()
    resolved = status.setdefault(agent, {}).setdefault(today, [])
    if idea in resolved:
        resolved.remove(idea)
    save_thought_status(status)
    return jsonify({'success': True})

@app.route('/thoughts/accept', methods=['POST'])
def thoughts_accept():
    data = request.json
    agent = data.get('agent')
    idea = data.get('idea')
    if not agent or agent not in ALL_AGENTS or not idea:
        return jsonify({'success': False, 'error': 'Missing agent or idea'}), 400

    status = load_thought_status()
    today = today_str()
    resolved = status.setdefault(agent, {}).setdefault(today, [])
    if idea not in resolved:
        resolved.append(idea)
    save_thought_status(status)
    return jsonify({'success': True})

@app.route('/personal/today', methods=['GET'])
def personal_today():
    kb = load_personal_knowledge_base()
    today = today_str()
    result = {}
    for agent in ALL_AGENTS:
        questions = get_today_personal_questions(agent)
        if not questions:
            continue
        agent_entries = kb.get(agent, [])
        items = []
        for q in questions:
            entry = next((e for e in agent_entries if e.get('date') == today and e.get('question') == q), None)
            items.append({
                'question': q,
                'answered': entry is not None,
                'answer': entry.get('answer') if entry else None
            })
        result[agent] = items
    return jsonify({'success': True, 'questions': result})

@app.route('/personal/answer', methods=['POST'])
def personal_answer():
    try:
        data = request.json
        agent = data.get('agent')
        question = data.get('question')
        answer = data.get('answer', '').strip()

        if not agent or agent not in ALL_AGENTS or not question or not answer:
            return jsonify({'success': False, 'error': 'Missing agent, question, or answer'}), 400

        today = today_str()

        with kb_answers_lock:
            kb = load_personal_knowledge_base()
            agent_entries = kb.setdefault(agent, [])

            existing = next((e for e in agent_entries if e.get('date') == today and e.get('question') == question), None)
            if existing:
                existing['answer'] = answer
                existing['merged'] = False
            else:
                agent_entries.append({'date': today, 'question': question, 'answer': answer, 'merged': False})

            save_personal_knowledge_base(kb)
        return jsonify({'success': True})
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/personal-thoughts/today', methods=['GET'])
def personal_thoughts_today():
    status = load_personal_thought_status()
    today = today_str()
    result = {}
    for agent in ALL_AGENTS:
        ideas = get_today_personal_ideas(agent)
        if not ideas:
            continue
        resolved = status.get(agent, {}).get(today, [])
        remaining = [idea for idea in ideas if idea not in resolved]
        if remaining:
            result[agent] = remaining
    return jsonify({'success': True, 'thoughts': result})

@app.route('/personal-thoughts/pass', methods=['POST'])
def personal_thoughts_pass():
    data = request.json
    agent = data.get('agent')
    idea = data.get('idea')
    if not agent or agent not in ALL_AGENTS or not idea:
        return jsonify({'success': False, 'error': 'Missing agent or idea'}), 400

    status = load_personal_thought_status()
    today = today_str()
    resolved = status.setdefault(agent, {}).setdefault(today, [])
    if idea not in resolved:
        resolved.append(idea)
    save_personal_thought_status(status)
    return jsonify({'success': True})

# Undoes a Pass - removes the idea from today's resolved list so it reappears
# as a pending suggestion. Used by the "Undo" button on the auto-logged chat
# record left behind when a suggestion is passed.
@app.route('/personal-thoughts/unpass', methods=['POST'])
def personal_thoughts_unpass():
    data = request.json
    agent = data.get('agent')
    idea = data.get('idea')
    if not agent or agent not in ALL_AGENTS or not idea:
        return jsonify({'success': False, 'error': 'Missing agent or idea'}), 400

    status = load_personal_thought_status()
    today = today_str()
    resolved = status.setdefault(agent, {}).setdefault(today, [])
    if idea in resolved:
        resolved.remove(idea)
    save_personal_thought_status(status)
    return jsonify({'success': True})

@app.route('/personal-thoughts/accept', methods=['POST'])
def personal_thoughts_accept():
    data = request.json
    agent = data.get('agent')
    idea = data.get('idea')
    if not agent or agent not in ALL_AGENTS or not idea:
        return jsonify({'success': False, 'error': 'Missing agent or idea'}), 400

    status = load_personal_thought_status()
    today = today_str()
    resolved = status.setdefault(agent, {}).setdefault(today, [])
    if idea not in resolved:
        resolved.append(idea)
    save_personal_thought_status(status)
    return jsonify({'success': True})

# How much conversation each message carries: the last HISTORY_VERBATIM_MESSAGES
# in full, plus a one-line-each digest of the HISTORY_DIGEST_MESSAGES before
# them (see build_history_digest) - so an agent keeps a sense of what came
# earlier without paying to re-read all of it every time. Attachments are only
# re-sent for the last HISTORY_ATTACHMENT_RECENT messages.
HISTORY_VERBATIM_MESSAGES = 10
HISTORY_DIGEST_MESSAGES = 10
HISTORY_ATTACHMENT_RECENT = 4


def build_history_digest(history, agent):
    older = (history or [])[:-HISTORY_VERBATIM_MESSAGES][-HISTORY_DIGEST_MESSAGES:]
    lines = []
    for entry in older:
        text = ' '.join((entry.get('text') or '').split())
        if not text:
            continue
        if entry.get('type') == 'agent':
            speaker = 'You' if not entry.get('agent') or entry.get('agent') == agent else entry['agent'].capitalize()
            limit = 110
        else:
            speaker, limit = 'Francis', 150
        lines.append(f"- {speaker}: {text[:limit]}{'...' if len(text) > limit else ''}")
    if not lines:
        return ''
    return (
        "\n\nEARLIER IN THIS CONVERSATION (older messages, heavily abbreviated - the recent messages "
        "follow in full below; ask if you need a detail that isn't here):\n" + "\n".join(lines)
    )


def build_claude_messages(history, agent, message, message_attachments=None):
    """Turn prior chat history into an alternating user/assistant message list for
    Claude, so replies have real conversational context instead of answering each
    message in isolation. This agent's own past messages become 'assistant' turns,
    everything else (the human, or a referral-context line from another agent)
    becomes a 'user' turn, with those other lines prefixed by their name so the
    model doesn't mistake them for the human speaking.

    A history entry that originally carried attachments keeps them here too -
    otherwise a later turn shows the model its own past analysis of a file with
    no file anywhere in its actual context, and it (correctly, from what it can
    see) concludes it was never sent one and disowns its own earlier answer.
    """
    turns = []
    recent = (history or [])[-HISTORY_VERBATIM_MESSAGES:]
    for position, entry in enumerate(recent):
        text = (entry.get('text') or '').strip()
        entry_attachments = entry.get('attachments') or []
        # A file attached a few messages back doesn't need to ride along in
        # full on every later turn - only a note that it was shared.
        if entry_attachments and position < len(recent) - HISTORY_ATTACHMENT_RECENT:
            names = ', '.join(str(a.get('name') or 'file') for a in entry_attachments if isinstance(a, dict))
            text = (text + '\n' if text else '') + f'[Earlier attachment(s): {names}]'
            entry_attachments = []
        if not text and not entry_attachments:
            continue
        if entry.get('type') == 'agent':
            if entry.get('agent') and entry.get('agent') != agent:
                role, text = 'user', f"[{entry['agent'].capitalize()}]: {text}"
            else:
                role = 'assistant'
        else:
            role = 'user'

        if entry_attachments:
            blocks = []
            if text:
                blocks.append({'type': 'text', 'text': text})
            blocks.extend(build_attachment_content_blocks(entry_attachments))
            turns.append({'role': role, 'content': blocks})
            continue

        if turns and turns[-1]['role'] == role and isinstance(turns[-1]['content'], str):
            turns[-1]['content'] += f"\n{text}"
        else:
            turns.append({'role': role, 'content': text})

    if turns and turns[-1]['role'] == 'user' and isinstance(turns[-1]['content'], str):
        turns[-1]['content'] += f"\n{message}"
    else:
        turns.append({'role': 'user', 'content': message})

    return turns

WORD_MIME_TYPES = {'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}
EXCEL_MIME_TYPES = {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
POWERPOINT_MIME_TYPES = {'application/vnd.openxmlformats-officedocument.presentationml.presentation'}

def extract_docx_text(file_bytes):
    doc = docx_lib.Document(io.BytesIO(file_bytes))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            parts.append(' | '.join(cell.text.strip() for cell in row.cells))
    return '\n'.join(parts) or '(no readable text found in this document)'

def extract_xlsx_text(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets_text = []
    for sheet in wb.worksheets:
        rows_text = []
        for row in sheet.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                rows_text.append(' | '.join('' if c is None else str(c) for c in row))
        if rows_text:
            sheets_text.append(f"Sheet: {sheet.title}\n" + '\n'.join(rows_text))
    return '\n\n'.join(sheets_text) or '(no data found in this spreadsheet)'

def extract_pptx_text(file_bytes):
    prs = pptx_lib.Presentation(io.BytesIO(file_bytes))
    slides_text = []
    for i, slide in enumerate(prs.slides, 1):
        lines = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                lines.append(shape.text_frame.text.strip())
            elif shape.has_table:
                for row in shape.table.rows:
                    lines.append(' | '.join(cell.text.strip() for cell in row.cells))
        if lines:
            slides_text.append(f"Slide {i}:\n" + '\n'.join(lines))
    return '\n\n'.join(slides_text) or '(no readable text found in this presentation)'


# Dependency-free HTML-to-text for an email's HTML body (used when a message
# has no plain-text part) - doesn't need to be a real HTML parser, just good
# enough to turn "<p>Hi <b>Francis</b>,</p>" into readable text for Claude.
def _html_to_text(html):
    text = re.sub(r'<(script|style)[^>]*>.*?</\1>', '', html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p>', '\n\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = html_lib.unescape(text)
    return re.sub(r'\n{3,}', '\n\n', text).strip()


# Dragging a real email out of desktop Outlook produces a virtual file whose
# name/browser-reported MIME type aren't reliable - Outlook usually names it
# after the subject line with no guarantee of a .msg/.eml extension, and
# Chromium often reports an empty file.type for extensions it doesn't
# recognize. Sniffing the actual bytes catches those cases: .msg is always an
# OLE2 compound file (fixed magic number), .eml is plain RFC 822 text
# identifiable by its header lines.
_OLE2_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
_EML_HEADER_MARKERS = ('from:', 'to:', 'subject:', 'date:', 'content-type:', 'mime-version:', 'return-path:', 'received:')

def _looks_like_msg_bytes(file_bytes):
    return file_bytes[:8] == _OLE2_MAGIC

def _looks_like_eml_bytes(file_bytes):
    head = file_bytes[:2000].decode('utf-8', errors='ignore').lower()
    return sum(1 for marker in _EML_HEADER_MARKERS if marker in head) >= 2

# .eml is the standard format most desktop/web mail clients export to - it's
# plain text (RFC 822), so Python's own email module reads it with no new
# dependency. The generic text-fallback in build_attachment_content_blocks
# would technically "work" on one of these too, but often garbles the body
# if it's MIME-multipart or base64/quoted-printable encoded - this parses it
# properly instead of hoping Claude can untangle raw MIME.
def extract_eml_text(file_bytes):
    parsed = email_lib.message_from_bytes(file_bytes, policy=email.policy.default)
    body = ''
    plain_part = parsed.get_body(preferencelist=('plain',))
    if plain_part is not None:
        body = plain_part.get_content()
    else:
        html_part = parsed.get_body(preferencelist=('html',))
        if html_part is not None:
            body = _html_to_text(html_part.get_content())
    lines = [
        f"From: {parsed.get('From', '')}",
        f"To: {parsed.get('To', '')}",
        f"Date: {parsed.get('Date', '')}",
        f"Subject: {parsed.get('Subject', '')}",
        "",
        (body or '').strip() or '(no readable body found in this email)'
    ]
    return '\n'.join(lines)


# .msg is Outlook's own binary format (an OLE compound file, not plain text) -
# what you get when dragging an email straight out of the Outlook desktop
# app. extract_msg already handles the RTF/HTML deencapsulation that a body
# can be stored in, so .body is preferred and .htmlBody is only a fallback.
def extract_msg_text(file_bytes):
    msg = extract_msg.Message(io.BytesIO(file_bytes))
    try:
        body = (msg.body or '').strip()
        if not body and msg.htmlBody:
            html_content = msg.htmlBody
            if isinstance(html_content, bytes):
                html_content = html_content.decode('utf-8', errors='replace')
            body = _html_to_text(html_content)
        lines = [
            f"From: {msg.sender or ''}",
            f"To: {msg.to or ''}",
            f"Date: {msg.date or ''}",
            f"Subject: {msg.subject or ''}",
            "",
            body or '(no readable body found in this email)'
        ]
        return '\n'.join(lines)
    finally:
        msg.close()

# HTML-body counterparts to extract_eml_text/extract_msg_text - used only by
# the preview endpoint (build_attachment_content_blocks sends Claude the
# plain-text version; markup adds nothing for the model to read and just
# burns tokens). Returns None when the email has no HTML part, which is
# common for plain-text-only emails.
def extract_eml_html(file_bytes):
    parsed = email_lib.message_from_bytes(file_bytes, policy=email.policy.default)
    html_part = parsed.get_body(preferencelist=('html',))
    return html_part.get_content() if html_part is not None else None

def extract_msg_html(file_bytes):
    msg = extract_msg.Message(io.BytesIO(file_bytes))
    try:
        html_content = msg.htmlBody
        if not html_content:
            return None
        if isinstance(html_content, bytes):
            html_content = html_content.decode('utf-8', errors='replace')
        return html_content
    finally:
        msg.close()

def extract_email_like_html(file_bytes, name, mime_type):
    lower_name = (name or '').lower()
    mime_type = mime_type or ''
    if mime_type in ('application/vnd.ms-outlook', 'application/x-msg') or lower_name.endswith('.msg') or _looks_like_msg_bytes(file_bytes):
        return extract_msg_html(file_bytes)
    if mime_type in ('message/rfc822', 'application/eml') or lower_name.endswith('.eml') or _looks_like_eml_bytes(file_bytes):
        return extract_eml_html(file_bytes)
    return None

# Shared by build_attachment_content_blocks (the /chat pipeline) and
# /extract-email-preview - name/mimeType are checked first (cheap, and correct
# when present), falling back to sniffing the bytes when they aren't.
# Returns None if this doesn't look like an email at all.
def extract_email_like_text(file_bytes, name, mime_type):
    lower_name = (name or '').lower()
    mime_type = mime_type or ''
    if mime_type in ('application/vnd.ms-outlook', 'application/x-msg') or lower_name.endswith('.msg') or _looks_like_msg_bytes(file_bytes):
        return extract_msg_text(file_bytes)
    if mime_type in ('message/rfc822', 'application/eml') or lower_name.endswith('.eml') or _looks_like_eml_bytes(file_bytes):
        return extract_eml_text(file_bytes)
    return None

@app.route('/extract-email-preview', methods=['POST'])
def extract_email_preview():
    try:
        data = request.json or {}
        name = data.get('name') or ''
        mime_type = data.get('mimeType') or ''
        file_bytes = base64.b64decode(data.get('data') or '')

        text = extract_email_like_text(file_bytes, name, mime_type)
        if text is None:
            return jsonify({'success': False, 'error': 'Unsupported file type for preview'}), 400

        try:
            html = extract_email_like_html(file_bytes, name, mime_type)
        except Exception:
            html = None

        return jsonify({'success': True, 'text': text, 'html': html})
    except Exception as e:
        print(f"Email preview extract error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

# Converts frontend-provided attachments (images, PDFs, Word, Excel, plain text)
# into Claude message content blocks. Images and PDFs go straight to Claude as
# native binary content (real vision / document understanding); Office formats
# have no native binary support in the API, so their text is extracted here on
# the server and sent as a labeled text block instead.
def build_attachment_content_blocks(attachments):
    blocks = []
    for att in (attachments or []):
        name = att.get('name', 'file')
        mime_type = att.get('mimeType', '') or ''
        data_b64 = att.get('data', '')
        if not data_b64:
            continue

        try:
            if mime_type.startswith('image/'):
                blocks.append({
                    'type': 'image',
                    'source': {'type': 'base64', 'media_type': mime_type, 'data': data_b64}
                })
                continue

            if mime_type == 'application/pdf':
                blocks.append({
                    'type': 'document',
                    'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': data_b64}
                })
                continue

            file_bytes = base64.b64decode(data_b64)
            lower_name = name.lower()

            if mime_type in WORD_MIME_TYPES or lower_name.endswith('.docx'):
                text = extract_docx_text(file_bytes)
            elif mime_type in EXCEL_MIME_TYPES or lower_name.endswith('.xlsx'):
                text = extract_xlsx_text(file_bytes)
            elif mime_type in POWERPOINT_MIME_TYPES or lower_name.endswith('.pptx'):
                text = extract_pptx_text(file_bytes)
            elif mime_type == 'application/msword' or lower_name.endswith('.doc'):
                text = "(Legacy .doc format can't be read directly - please save as .docx and re-attach.)"
            elif mime_type == 'application/vnd.ms-excel' or lower_name.endswith('.xls'):
                text = "(Legacy .xls format can't be read directly - please save as .xlsx and re-attach.)"
            elif mime_type == 'application/vnd.ms-powerpoint' or lower_name.endswith('.ppt'):
                text = "(Legacy .ppt format can't be read directly - please save as .pptx and re-attach.)"
            else:
                email_text = extract_email_like_text(file_bytes, name, mime_type)
                if email_text is not None:
                    text = email_text
                else:
                    # Plain text and anything else decodable as text (.txt, .csv, .md, etc.)
                    text = file_bytes.decode('utf-8', errors='replace')

            blocks.append({'type': 'text', 'text': f"[Attached file: {name}]\n{text.strip()}"})
        except Exception as e:
            blocks.append({'type': 'text', 'text': f"[Attached file: {name} - couldn't be read: {str(e)}]"})

    return blocks

# --- File creation -----------------------------------------------------
# An agent that hands over a real deliverable (not just a chat answer) calls
# create_file with plain-text content in one of a few light conventions
# (documented in CREATE_FILE_TOOL below); these functions turn that into an
# actual downloadable file in the requested format.

FILE_TYPE_MIME = {
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'pdf': 'application/pdf',
    'txt': 'text/plain',
    'csv': 'text/csv',
    'html': 'text/html',
}

# A small set of named color schemes an agent can pick between when generating
# (or regenerating, with the same content, in response to "I don't like the
# colors" / "try a different look") a file. Each one supplies a dark "primary"
# used for headings and dark backgrounds, and a bright "accent" used for rules,
# bars, and highlights. Body text/gray stays constant across themes since a
# neutral gray reads fine against any of them.
THEMES = {
    'navy_gold':      {'primary': (0x1F, 0x3A, 0x5F), 'accent': (0xC9, 0x9A, 0x2E)},
    'charcoal_teal':  {'primary': (0x26, 0x2B, 0x2E), 'accent': (0x1F, 0x9E, 0x8B)},
    'burgundy_slate': {'primary': (0x5C, 0x1A, 0x2B), 'accent': (0xB0, 0x8D, 0x57)},
    'forest_emerald': {'primary': (0x1B, 0x3A, 0x2B), 'accent': (0xC9, 0xA9, 0x4E)},
    'slate_blue':     {'primary': (0x2E, 0x3D, 0x4F), 'accent': (0x7E, 0xB6, 0xD9)},
}
DEFAULT_THEME = 'navy_gold'
BRAND_GRAY = (0x33, 0x33, 0x33)
BRAND_LIGHT_GRAY = (0xF2, 0xF2, 0xF2)

def parse_hex_color(hex_str):
    if not hex_str:
        return None
    s = hex_str.strip().lstrip('#')
    if len(s) != 6 or not re.fullmatch(r'[0-9a-fA-F]{6}', s):
        return None
    return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))

# primary_hex/accent_hex (e.g. "#2E5E4E") let an agent match any color Francis
# names exactly, overriding whichever preset theme is passed. Falls back to the
# named theme (or the default) for whichever of the two isn't a valid hex code.
def resolve_theme(theme_name, primary_hex=None, accent_hex=None):
    theme = THEMES.get((theme_name or '').strip().lower(), THEMES[DEFAULT_THEME])
    primary_rgb = parse_hex_color(primary_hex) or theme['primary']
    accent_rgb = parse_hex_color(accent_hex) or theme['accent']
    return primary_rgb, accent_rgb

# Claude writes in markdown by habit even when told to write plain text for a
# generated file - **bold**, `code`, and stray leading #'s leak through as
# literal characters in Word/PDF/PowerPoint output otherwise (none of those
# formats render markdown syntax). Strip it so the file looks finished, not
# like raw markdown pasted into a document.
def strip_inline_markdown(text):
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'(?<!\*)\*([^*\n]+?)\*(?!\*)', r'\1', text)
    text = re.sub(r'`([^`]+?)`', r'\1', text)
    return text

def strip_leading_heading_marks(text):
    return re.sub(r'^#{1,6}\s*', '', text)

def _docx_add_bottom_border(paragraph, hex_color):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = DocxOxmlElement('w:pBdr')
    bottom = DocxOxmlElement('w:bottom')
    bottom.set(docx_qn('w:val'), 'single')
    bottom.set(docx_qn('w:sz'), '18')
    bottom.set(docx_qn('w:space'), '6')
    bottom.set(docx_qn('w:color'), hex_color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)

# Colors every heading navy and underlines the document's main title with a
# gold rule, so a generated doc reads as a designed deliverable rather than
# Word's bare default black-on-white styling.
def create_docx_bytes(content, theme=DEFAULT_THEME, primary_color=None, accent_color=None):
    primary_rgb, accent_rgb = resolve_theme(theme, primary_color, accent_color)
    doc = docx_lib.Document()
    navy = DocxRGBColor(*primary_rgb)
    gold_hex = '{:02X}{:02X}{:02X}'.format(*accent_rgb)
    first_h1_done = False

    for raw_line in content.split('\n'):
        stripped = raw_line.strip()
        if not stripped:
            continue
        if stripped.startswith('### '):
            heading = doc.add_heading(strip_inline_markdown(stripped[4:].strip()), level=3)
            for run in heading.runs:
                run.font.color.rgb = navy
        elif stripped.startswith('## '):
            heading = doc.add_heading(strip_inline_markdown(stripped[3:].strip()), level=2)
            for run in heading.runs:
                run.font.color.rgb = navy
        elif stripped.startswith('# '):
            heading = doc.add_heading(strip_inline_markdown(stripped[2:].strip()), level=1)
            for run in heading.runs:
                run.font.color.rgb = navy
                run.font.size = DocxPt(24)
            if not first_h1_done:
                _docx_add_bottom_border(heading, gold_hex)
                heading.paragraph_format.space_after = DocxPt(14)
                first_h1_done = True
        elif stripped.startswith('- ') or stripped.startswith('* '):
            doc.add_paragraph(strip_inline_markdown(stripped[2:].strip()), style='List Bullet')
        else:
            doc.add_paragraph(strip_inline_markdown(stripped))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def parse_cell_value(value):
    value = value.strip()
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        return value

def create_xlsx_bytes(content, theme=DEFAULT_THEME, primary_color=None, accent_color=None):
    primary_rgb, accent_rgb = resolve_theme(theme, primary_color, accent_color)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    rows = [line for line in content.split('\n') if line.strip()]

    navy_hex = '{:02X}{:02X}{:02X}'.format(*primary_rgb)
    header_font = XlsxFont(bold=True, color='FFFFFF')
    header_fill = XlsxFill(start_color=navy_hex, end_color=navy_hex, fill_type='solid')

    for r, row_line in enumerate(rows, 1):
        cells = [c.strip() for c in row_line.split('|')]
        for c, raw_value in enumerate(cells, 1):
            value = raw_value if r == 1 else parse_cell_value(raw_value)
            cell = ws.cell(row=r, column=c, value=value)
            if r == 1:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = XlsxAlignment(vertical='center')
            else:
                cell.alignment = XlsxAlignment(vertical='center')
    if rows:
        for c in range(1, len(rows[0].split('|')) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 24
        ws.row_dimensions[1].height = 20
        ws.freeze_panes = 'A2'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def create_csv_bytes(content):
    rows = [line for line in content.split('\n') if line.strip()]
    output = io.StringIO()
    writer = csv_lib.writer(output)
    for row_line in rows:
        writer.writerow([c.strip() for c in row_line.split('|')])
    return output.getvalue().encode('utf-8')

def clean_pptx_line(line):
    return strip_inline_markdown(strip_leading_heading_marks(line)).strip()

def _pptx_add_bullets(text_frame, lines, size_pt, color_rgb):
    text_frame.word_wrap = True
    for i, line in enumerate(lines):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        p.text = f"●   {line}"
        p.font.size = Pt(size_pt)
        p.font.color.rgb = PptxRGBColor(*color_rgb)
        p.space_after = Pt(10)

# Builds an actual designed deck (a real title slide plus a consistent color
# scheme and accent bar on every content slide) instead of dumping text onto
# python-pptx's bare default template - a generated deck is meant to be handed
# straight to a client, not to look like an unstyled draft.
def create_pptx_bytes(content, theme=DEFAULT_THEME, primary_color=None, accent_color=None):
    primary_rgb, accent_rgb = resolve_theme(theme, primary_color, accent_color)
    prs = pptx_lib.Presentation()
    blank_layout = prs.slide_layouts[6]  # fully blank - full manual control over design
    slide_blocks = [b for b in content.split('\n---\n')]

    for idx, slide_content in enumerate(slide_blocks):
        lines = [
            clean_pptx_line(l) for l in slide_content.split('\n')
            if l.strip() and l.strip() != '---'
        ]
        lines = [l for l in lines if l]
        if not lines:
            continue

        title_text = lines[0]
        body_lines = [clean_pptx_line(l.lstrip('-* ')) for l in lines[1:]]
        body_lines = [l for l in body_lines if l]

        slide = prs.slides.add_slide(blank_layout)
        slide.background.fill.solid()

        if idx == 0:
            # Full-bleed title slide with an accent rule and subtitle.
            slide.background.fill.fore_color.rgb = PptxRGBColor(*primary_rgb)

            accent_rule = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.55), Inches(1.4), Inches(0.06)
            )
            accent_rule.fill.solid()
            accent_rule.fill.fore_color.rgb = PptxRGBColor(*accent_rgb)
            accent_rule.line.fill.background()
            accent_rule.shadow.inherit = False

            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.75), prs.slide_width - Inches(1.6), Inches(1.6))
            tf = title_box.text_frame
            tf.word_wrap = True
            tf.text = title_text
            tf.paragraphs[0].font.size = Pt(40)
            tf.paragraphs[0].font.bold = True
            tf.paragraphs[0].font.color.rgb = PptxRGBColor(0xFF, 0xFF, 0xFF)

            if body_lines:
                subtitle_box = slide.shapes.add_textbox(Inches(0.8), Inches(4.3), prs.slide_width - Inches(1.6), Inches(1.8))
                stf = subtitle_box.text_frame
                stf.word_wrap = True
                stf.text = body_lines[0]
                stf.paragraphs[0].font.size = Pt(18)
                stf.paragraphs[0].font.color.rgb = PptxRGBColor(*accent_rgb)
                for extra in body_lines[1:]:
                    p = stf.add_paragraph()
                    p.text = extra
                    p.font.size = Pt(15)
                    p.font.color.rgb = PptxRGBColor(0xE0, 0xE0, 0xE0)
            continue

        # Content slide: white background, primary-colored title, accent top bar,
        # gray bulleted body - the same theme colors on every slide.
        slide.background.fill.fore_color.rgb = PptxRGBColor(0xFF, 0xFF, 0xFF)

        accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.18))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = PptxRGBColor(*accent_rgb)
        accent_bar.line.fill.background()
        accent_bar.shadow.inherit = False

        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), prs.slide_width - Inches(1.2), Inches(1.0))
        ttf = title_box.text_frame
        ttf.word_wrap = True
        ttf.text = title_text
        ttf.paragraphs[0].font.size = Pt(28)
        ttf.paragraphs[0].font.bold = True
        ttf.paragraphs[0].font.color.rgb = PptxRGBColor(*primary_rgb)

        if body_lines:
            body_box = slide.shapes.add_textbox(
                Inches(0.7), Inches(1.5), prs.slide_width - Inches(1.4), prs.slide_height - Inches(1.9)
            )
            _pptx_add_bullets(body_box.text_frame, body_lines, 18, BRAND_GRAY)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# fpdf2's built-in fonts only cover Latin-1, but Claude's own writing style
# leans on em dashes and curly quotes constantly - using the system's real
# Arial (present on any Windows install) instead gives full Unicode support;
# ASCII-sanitizing is only the fallback if that font can't be found.
def find_windows_font(filename):
    path = os.path.join('C:\\Windows\\Fonts', filename)
    return path if os.path.exists(path) else None

def sanitize_for_core_font(text):
    replacements = {
        '\u2014': '-', '\u2013': '-', '\u2018': "'", '\u2019': "'",
        '\u201c': '"', '\u201d': '"', '\u2026': '...', '\u2022': '*'
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.encode('latin-1', errors='replace').decode('latin-1')

def create_pdf_bytes(content, theme=DEFAULT_THEME, primary_color=None, accent_color=None):
    primary_rgb, accent_rgb = resolve_theme(theme, primary_color, accent_color)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    regular_font = find_windows_font('arial.ttf')
    bold_font = find_windows_font('arialbd.ttf') or regular_font

    if regular_font:
        pdf.add_font('Body', '', regular_font)
        pdf.add_font('Body', 'B', bold_font)
        base_font = 'Body'
    else:
        content = sanitize_for_core_font(content)
        base_font = 'Helvetica'

    pdf.set_font(base_font, '', 11)
    pdf.set_text_color(*BRAND_GRAY)
    first_h1_done = False
    for raw_line in content.split('\n'):
        stripped = raw_line.strip()
        if not stripped:
            pdf.ln(4)
            continue
        # multi_cell(w=0, ...) sizes itself from the CURRENT x position, not the
        # left margin - without resetting x first, it eats into its own
        # available width on every call until there's none left.
        pdf.set_x(pdf.l_margin)
        if stripped.startswith('### '):
            pdf.set_font(base_font, 'B', 13)
            pdf.set_text_color(*primary_rgb)
            pdf.multi_cell(0, 8, strip_inline_markdown(stripped[4:].strip()))
            pdf.set_font(base_font, '', 11)
            pdf.set_text_color(*BRAND_GRAY)
        elif stripped.startswith('## '):
            pdf.set_font(base_font, 'B', 15)
            pdf.set_text_color(*primary_rgb)
            pdf.multi_cell(0, 9, strip_inline_markdown(stripped[3:].strip()))
            pdf.set_font(base_font, '', 11)
            pdf.set_text_color(*BRAND_GRAY)
        elif stripped.startswith('# '):
            pdf.set_font(base_font, 'B', 22)
            pdf.set_text_color(*primary_rgb)
            pdf.multi_cell(0, 12, strip_inline_markdown(stripped[2:].strip()))
            if not first_h1_done:
                # An accent rule under the document's main title - a small but
                # real branding touch instead of a bare heading on white.
                pdf.set_draw_color(*accent_rgb)
                pdf.set_line_width(0.8)
                pdf.line(pdf.l_margin, pdf.get_y() + 1, pdf.w - pdf.r_margin, pdf.get_y() + 1)
                pdf.ln(6)
                first_h1_done = True
            pdf.set_font(base_font, '', 11)
            pdf.set_text_color(*BRAND_GRAY)
        elif stripped.startswith('- ') or stripped.startswith('* '):
            pdf.multi_cell(0, 7, f"-  {strip_inline_markdown(stripped[2:].strip())}")
        else:
            pdf.multi_cell(0, 7, strip_inline_markdown(stripped))

    return bytes(pdf.output())

def generate_file_bytes(file_type, content, theme=DEFAULT_THEME, primary_color=None, accent_color=None):
    if file_type == 'docx':
        return create_docx_bytes(content, theme, primary_color, accent_color)
    if file_type == 'xlsx':
        return create_xlsx_bytes(content, theme, primary_color, accent_color)
    if file_type == 'pptx':
        return create_pptx_bytes(content, theme, primary_color, accent_color)
    if file_type == 'pdf':
        return create_pdf_bytes(content, theme, primary_color, accent_color)
    if file_type == 'csv':
        return create_csv_bytes(content)
    if file_type in ('txt', 'html'):
        return content.encode('utf-8')
    raise ValueError(f"Unsupported file_type: {file_type}")

# Rebuilds a file's downloadable bytes from its (possibly hand-edited) text, so
# a download always matches what the Workspace shows after Francis edits it
# there. Same generator create_file uses - no model involved.
@app.route('/files/regenerate', methods=['POST'])
def files_regenerate():
    try:
        data = request.json or {}
        file_type = str(data.get('file_type') or '').strip().lower()
        content = str(data.get('content') or '')
        if file_type not in FILE_TYPE_MIME or not content.strip():
            return jsonify({'success': False, 'error': 'Unsupported file type or empty content'}), 400
        if len(content) > 400000:
            return jsonify({'success': False, 'error': 'That file is too large to rebuild'}), 400
        theme = str(data.get('theme') or '').strip().lower() or DEFAULT_THEME
        raw = generate_file_bytes(file_type, content, theme, str(data.get('primary_color') or '').strip() or None,
                                  str(data.get('accent_color') or '').strip() or None)
        return jsonify({'success': True, 'data': base64.b64encode(raw).decode('ascii')})
    except Exception as e:
        print(f"File regenerate error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# Acted on by the frontend: renders as a downloadable file chip on the agent's
# message instead of (or alongside) plain chat text.
CREATE_FILE_TOOL = {
    "name": "create_file",
    "description": (
        "Only available while you're running a task - every file belongs to a task. Call this when "
        "you've prepared something Francis asked for as an actual file he can download "
        "and use directly - a document, spreadsheet, presentation, PDF, or web page/mockup - rather than "
        "pasting the content into chat. Use it when Francis asks for something 'as a doc/Word file/"
        "spreadsheet/Excel/PDF/deck/PowerPoint/HTML page/mockup', or when handing over a finished "
        "deliverable (a report, a template, a workbook, a visual preview) that naturally belongs in a "
        "real file. Don't use this for short answers or anything that reads fine as a normal chat "
        "message - only when a downloadable file is genuinely what's being asked for, and never claim "
        "in your reply that you attached or generated a file unless you actually called this tool in the "
        "same turn - a description of what a file would contain is not a file. Always pair this with a "
        "short reply of your own describing what you made - never call it as your only output. This tool "
        "always builds the file fresh from the content and theme you pass in - there's no way to edit a "
        "previously generated file in place. So if Francis asks to revise, tweak, redo, or change "
        "anything about a file you already made (the wording, a section, the color scheme, the whole "
        "look), just call create_file again with the updated content and/or a different theme - that "
        "regenerates the whole file with the changes applied. If he asked for the SAME content as more "
        "than one file type (e.g. 'a PDF and a Word version', or 'a PDF, a Word doc, and a PowerPoint'), "
        "call create_file once per format, all as separate calls - every file you produce gets attached, "
        "not just the last one."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "filename": {
                "type": "string",
                "description": "Filename without extension, e.g. \"Client Engagement Letter\"."
            },
            "file_type": {
                "type": "string",
                "enum": ["docx", "xlsx", "pptx", "pdf", "txt", "csv", "html"]
            },
            "theme": {
                "type": "string",
                "enum": ["navy_gold", "charcoal_teal", "burgundy_slate", "forest_emerald", "slate_blue"],
                "description": (
                    "Preset color scheme for docx/xlsx/pptx/pdf (ignored for txt/csv/html, which have no "
                    "built-in styling - an html file's look comes entirely from the CSS you write into its "
                    "content instead). Defaults to navy_gold if omitted. Used as-is if Francis just wants "
                    "'a different look', or as the fallback for whichever of primary_color/accent_color "
                    "below is left blank when he names a specific color."
                )
            },
            "primary_color": {
                "type": "string",
                "description": (
                    "Optional hex color (e.g. \"#2E5E4E\") to use instead of the theme's preset primary "
                    "color, for headings/dark backgrounds. Set this when Francis names or describes a "
                    "specific color he wants (e.g. \"make it forest green\", \"use our brand color "
                    "#003366\") rather than just picking a theme."
                )
            },
            "accent_color": {
                "type": "string",
                "description": "Optional hex color (e.g. \"#C9A94E\") for the accent rules/bars/highlights, same rules as primary_color."
            },
            "content": {
                "type": "string",
                "description": (
                    "For every file_type except html, PLAIN TEXT using ONLY these conventions - no other "
                    "markdown syntax at all (no **bold**, no backtick code, no stray # outside the heading "
                    "rule below), since none of these file formats render markdown and it will show up as "
                    "literal asterisks/hashes in the finished file: "
                    "for docx/pdf/txt - one line per paragraph, a line starting with \"# \" is a top-level "
                    "heading (\"## \"/\"### \" for smaller headings), and a line starting with \"- \" is a "
                    "bullet point, otherwise just write the words with no special characters for emphasis; "
                    "for xlsx/csv - one row per line, cells separated by \"|\", first row is the header, "
                    "plain numbers and text only in cells; for pptx - slides separated by a line containing "
                    "only \"---\", each slide's first line is its title (plain text, no \"#\") and the "
                    "remaining lines are its bullet points (plain text, no \"-\" needed - one point per line). "
                    "For file_type html, write a COMPLETE, self-contained HTML document instead (starting "
                    "with <!DOCTYPE html>, including <html>/<head>/<body>) with any CSS inline in a <style> "
                    "tag - no external stylesheets, fonts, or scripts, since the preview renders it "
                    "sandboxed with those blocked. Use this for mockups, color/design previews, or any "
                    "other visual page Francis wants to look at rather than read as a document."
                )
            }
        },
        "required": ["filename", "file_type", "content"]
    }
}

# Lets an agent explicitly say "I'm not actually done" - e.g. hitting its
# web_search cap partway through deep research - instead of that just
# reading as a normal finished reply. Only matters for a Started task (see
# task_paused handling in /chat and pauseProjectTask on the frontend); for an
# ordinary chat message the reply still shows normally either way.
PAUSE_TASK_TOOL = {
    "name": "pause_task",
    "description": (
        "Call this INSTEAD of just replying normally when you're working a background task (one "
        "Francis started from his Workspace) and genuinely cannot finish it in this turn - most "
        "commonly because you used up your available web searches partway through deep research and "
        "need to continue in a follow-up, but also if the task turns out to need something else "
        "(missing information, a decision from Francis) before you can complete it. Don't call this "
        "for a task you've actually completed, and don't call it outside of a Started task - for a "
        "normal quick question, just answer directly."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "reason": {
                "type": "string",
                "description": (
                    "A short, friendly note explaining why you're pausing and what's left - this is "
                    "shown to Francis directly as your reply, so write it that way (e.g. \"I've covered "
                    "eligibility and benefit calculation so far, but used up my searches before getting "
                    "to Medicare and covered services - click Start again and I'll pick up there.\")"
                )
            }
        },
        "required": ["reason"]
    }
}

# Triage tool for /classify-task-needed below - the prompt requires the model to
# call it, so the call returns exactly this shape instead of free text to parse
# (Sonnet 5.5 doesn't allow forcing a specific tool via tool_choice).
CLASSIFY_TASK_TOOL = {
    "name": "classify",
    "description": "Classify whether this request needs a background task.",
    "input_schema": {
        "type": "object",
        "properties": {
            "needs_task": {
                "type": "boolean",
                "description": (
                    "True if answering this well genuinely requires deep research (several web "
                    "searches to dig through a topic or a whole site) or creating a downloadable file "
                    "(a Word doc, spreadsheet, presentation, PDF, or HTML page/mockup) - or doing real "
                    "work WITH attached files (analyzing them into a report, editing, converting, "
                    "filling in, reconciling, building something from them) - real work that "
                    "takes meaningful time. False for anything answerable directly in a normal quick "
                    "reply: short questions, chit-chat, a quick opinion or explanation, or ordinary task/"
                    "project management (creating, editing, or discussing a task)."
                )
            },
            "task_name": {
                "type": "string",
                "description": "A short (4-8 word) task name, only when needs_task is true."
            }
        },
        "required": ["needs_task"]
    }
}


# --- Social accounts and publishing ---------------------------------------
# Sasha drafts posts; nothing is ever published by an agent. Connecting an
# account happens in Settings > Agents > Sasha (OAuth against the platform's
# own login page), and a post only goes out when Francis presses Publish on
# the review screen, which calls /social/publish with confirm=true. No chat
# tool can reach any of this, so nothing an agent writes can post itself.
#
# Each platform is a small provider: is it configured (developer-app keys in
# the server environment), the login URL, finishing the login, and publishing
# text. A "practice" provider connects instantly and only records what would
# have been posted, so the whole flow can be rehearsed without any account.
SOCIAL_FILE = _data_path('social_accounts.json')
SOCIAL_LOG_FILE = _data_path('social_publish_log.jsonl')
LINKEDIN_API_VERSION = os.getenv('LINKEDIN_API_VERSION', '202609')
FACEBOOK_GRAPH_VERSION = os.getenv('FACEBOOK_GRAPH_VERSION', 'v25.0')
SOCIAL_LABELS = {'linkedin': 'LinkedIn', 'facebook': 'Facebook Page', 'practice': 'Practice mode'}
SOCIAL_MAX_LENGTH = {'linkedin': 3000, 'facebook': 20000, 'practice': 20000}
social_lock = threading.Lock()
_social_recent_publishes = {}   # (publish id, platform) -> result, so a double click never double-posts


class SocialError(Exception):
    pass


def _social_load():
    if os.path.exists(SOCIAL_FILE):
        try:
            with open(SOCIAL_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _social_save(data):
    with open(SOCIAL_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def _social_error_text(status, body):
    message = ''
    try:
        parsed = json.loads(body)
        err = parsed.get('error') if isinstance(parsed, dict) else None
        if isinstance(err, dict):
            message = err.get('message') or ''
        elif isinstance(err, str):
            message = parsed.get('error_description') or err
        if not message and isinstance(parsed, dict):
            message = parsed.get('message') or ''
    except ValueError:
        pass
    message = message or body[:200]
    message = re.sub(r'(access_token|client_secret|fb_exchange_token)=[^&\s"]+', r'\1=***', message)
    return f"{message.strip()[:300]} (HTTP {status})"


def _social_request(method, url, headers=None, form=None, json_body=None, timeout=25):
    data = None
    hdrs = dict(headers or {})
    if form is not None:
        data = urllib.parse.urlencode(form).encode('utf-8')
        hdrs.setdefault('Content-Type', 'application/x-www-form-urlencoded')
    elif json_body is not None:
        data = json.dumps(json_body).encode('utf-8')
        hdrs.setdefault('Content-Type', 'application/json')
    request_obj = urllib.request.Request(url, data=data, headers=hdrs, method=method)
    try:
        with urllib.request.urlopen(request_obj, timeout=timeout) as resp:
            body = resp.read().decode('utf-8', 'replace')
            status = resp.status
            resp_headers = {k.lower(): v for k, v in resp.headers.items()}
    except urllib.error.HTTPError as e:
        raise SocialError(_social_error_text(e.code, e.read().decode('utf-8', 'replace')))
    except urllib.error.URLError as e:
        raise SocialError(f"Couldn't reach the service: {e.reason}")
    parsed = {}
    if body.strip().startswith(('{', '[')):
        try:
            parsed = json.loads(body)
        except ValueError:
            parsed = {}
    return status, resp_headers, parsed


def _social_redirect_uri(platform):
    base = (os.getenv('PUBLIC_BASE_URL') or request.url_root).rstrip('/')
    return f"{base}/social/callback/{platform}"


def _social_expires_at(seconds):
    return (now_local() + timedelta(seconds=int(seconds))).isoformat()


# -- LinkedIn (posts to the connected member's own profile) ------------------
def _linkedin_configured():
    return bool(os.getenv('LINKEDIN_CLIENT_ID') and os.getenv('LINKEDIN_CLIENT_SECRET'))


def _linkedin_auth_url(state, redirect_uri):
    return 'https://www.linkedin.com/oauth/v2/authorization?' + urllib.parse.urlencode({
        'response_type': 'code', 'client_id': os.getenv('LINKEDIN_CLIENT_ID'), 'redirect_uri': redirect_uri,
        'state': state, 'scope': 'openid profile w_member_social'
    })


def _linkedin_finish(code, redirect_uri):
    _, _, token = _social_request('POST', 'https://www.linkedin.com/oauth/v2/accessToken', form={
        'grant_type': 'authorization_code', 'code': code, 'redirect_uri': redirect_uri,
        'client_id': os.getenv('LINKEDIN_CLIENT_ID'), 'client_secret': os.getenv('LINKEDIN_CLIENT_SECRET')
    })
    access = token.get('access_token')
    if not access:
        raise SocialError('LinkedIn did not return an access token.')
    _, _, info = _social_request('GET', 'https://api.linkedin.com/v2/userinfo', headers={'Authorization': f'Bearer {access}'})
    if not info.get('sub'):
        raise SocialError("Couldn't read your LinkedIn profile.")
    return {
        'accessToken': access,
        # Self-serve LinkedIn tokens last about 60 days and cannot be refreshed.
        'expiresAt': _social_expires_at(token.get('expires_in') or 5184000),
        'accountName': info.get('name') or 'LinkedIn member',
        'authorUrn': f"urn:li:person:{info['sub']}"
    }


# LinkedIn's "little" text format treats these characters as syntax; unescaped
# they can swallow or cut off the rest of a post. #tags are turned into real
# hashtags with LinkedIn's hashtag template instead.
_LINKEDIN_RESERVED = re.compile(r'([\\|{}@\[\]()<>#*_~])')


def _linkedin_commentary(text):
    out = []
    pos = 0
    for match in re.finditer(r'(?<![\w\\])#([A-Za-z][A-Za-z0-9]*)', text):
        out.append(_LINKEDIN_RESERVED.sub(r'\\\1', text[pos:match.start()]))
        out.append('{hashtag|\\#|' + match.group(1) + '}')
        pos = match.end()
    out.append(_LINKEDIN_RESERVED.sub(r'\\\1', text[pos:]))
    return ''.join(out)


def _linkedin_publish(account, text):
    if account.get('expiresAt') and datetime.fromisoformat(account['expiresAt']) <= now_local():
        raise SocialError('The LinkedIn connection has expired - reconnect it in Settings > Agents > Sasha.')
    _, headers, _ = _social_request('POST', 'https://api.linkedin.com/rest/posts', headers={
        'Authorization': f"Bearer {account['accessToken']}",
        'Linkedin-Version': LINKEDIN_API_VERSION,
        'X-Restli-Protocol-Version': '2.0.0'
    }, json_body={
        'author': account['authorUrn'],
        'commentary': _linkedin_commentary(text),
        'visibility': 'PUBLIC',
        'distribution': {'feedDistribution': 'MAIN_FEED', 'targetEntities': [], 'thirdPartyDistributionChannels': []},
        'lifecycleState': 'PUBLISHED',
        'isReshareDisabledByAuthor': False
    })
    post_id = headers.get('x-restli-id') or ''
    return {'id': post_id, 'url': f"https://www.linkedin.com/feed/update/{post_id}/" if post_id else None}


# -- Facebook Pages ------------------------------------------------------------
def _facebook_configured():
    return bool(os.getenv('META_APP_ID') and os.getenv('META_APP_SECRET'))


def _facebook_auth_url(state, redirect_uri):
    return f"https://www.facebook.com/{FACEBOOK_GRAPH_VERSION}/dialog/oauth?" + urllib.parse.urlencode({
        'client_id': os.getenv('META_APP_ID'), 'redirect_uri': redirect_uri, 'state': state,
        'scope': 'pages_show_list,pages_manage_posts,pages_read_engagement'
    })


def _facebook_finish(code, redirect_uri):
    graph = f"https://graph.facebook.com/{FACEBOOK_GRAPH_VERSION}"
    app_creds = {'client_id': os.getenv('META_APP_ID'), 'client_secret': os.getenv('META_APP_SECRET')}
    _, _, short = _social_request('GET', f"{graph}/oauth/access_token?" + urllib.parse.urlencode(
        dict(app_creds, redirect_uri=redirect_uri, code=code)))
    if not short.get('access_token'):
        raise SocialError('Facebook did not return an access token.')
    _, _, long_lived = _social_request('GET', f"{graph}/oauth/access_token?" + urllib.parse.urlencode(
        dict(app_creds, grant_type='fb_exchange_token', fb_exchange_token=short['access_token'])))
    user_token = long_lived.get('access_token') or short['access_token']
    # Page tokens fetched with a long-lived user token don't expire.
    _, _, accounts = _social_request('GET', f"{graph}/me/accounts?" + urllib.parse.urlencode(
        {'fields': 'id,name,access_token', 'access_token': user_token}))
    pages = [{'id': p['id'], 'name': p.get('name') or p['id'], 'accessToken': p['access_token']}
             for p in (accounts.get('data') or []) if p.get('id') and p.get('access_token')]
    if not pages:
        raise SocialError("No Facebook Pages came back. Log in with an account that manages a Page and allow access to it.")
    return {'pages': pages, 'pageId': pages[0]['id'], 'accountName': pages[0]['name'], 'expiresAt': None}


def _facebook_publish(account, text):
    page = next((p for p in account.get('pages', []) if p['id'] == account.get('pageId')), None)
    if not page:
        raise SocialError('Choose which Facebook Page to post to in Settings > Agents > Sasha.')
    _, _, result = _social_request('POST', f"https://graph.facebook.com/{FACEBOOK_GRAPH_VERSION}/{page['id']}/feed",
                                   form={'message': text, 'access_token': page['accessToken']})
    post_id = result.get('id') or ''
    return {'id': post_id, 'url': f"https://www.facebook.com/{post_id}" if post_id else None}


# -- Practice mode (nothing leaves the app) ---------------------------------------
def _practice_publish(account, text):
    return {'id': f"practice-{uuid.uuid4().hex[:8]}", 'url': None, 'practice': True}


SOCIAL_PROVIDERS = {
    'linkedin': {
        'configured': _linkedin_configured, 'auth_url': _linkedin_auth_url, 'finish': _linkedin_finish, 'publish': _linkedin_publish,
        'setup': [
            "Create an app at linkedin.com/developers/apps (it asks for a LinkedIn Page to attach it to).",
            "On the app's Products tab, add \"Share on LinkedIn\" and \"Sign In with LinkedIn using OpenID Connect\".",
            "On the Auth tab, add the redirect URL shown below.",
            "Put the Client ID and Client Secret in the server settings LINKEDIN_CLIENT_ID and LINKEDIN_CLIENT_SECRET, then restart the app."
        ]
    },
    'facebook': {
        'configured': _facebook_configured, 'auth_url': _facebook_auth_url, 'finish': _facebook_finish, 'publish': _facebook_publish,
        'setup': [
            "At developers.facebook.com, create an app and add the Facebook Login product.",
            "Add the redirect URL shown below under Valid OAuth Redirect URIs, and make sure you are an admin or tester of the app.",
            "The app asks for the permissions pages_show_list, pages_manage_posts and pages_read_engagement. In development mode these work for Pages you manage.",
            "Put the App ID and App Secret in the server settings META_APP_ID and META_APP_SECRET, then restart the app."
        ]
    },
    'practice': {'configured': lambda: True, 'publish': _practice_publish, 'setup': []}
}


def _social_public_status():
    saved = _social_load()
    platforms = []
    for pid, provider in SOCIAL_PROVIDERS.items():
        account = saved.get(pid) or {}
        connected = bool(account)
        expires_at = account.get('expiresAt')
        expired = bool(expires_at and datetime.fromisoformat(expires_at) <= now_local())
        entry = {
            'id': pid, 'label': SOCIAL_LABELS[pid], 'configured': provider['configured'](),
            'connected': connected, 'accountName': account.get('accountName'),
            'expiresAt': expires_at, 'expired': expired,
            'maxLength': SOCIAL_MAX_LENGTH[pid], 'setup': provider['setup'],
            'redirectUri': _social_redirect_uri(pid) if 'auth_url' in provider else None
        }
        if pid == 'facebook' and connected:
            entry['pages'] = [{'id': p['id'], 'name': p['name']} for p in account.get('pages', [])]
            entry['pageId'] = account.get('pageId')
        platforms.append(entry)
    return platforms


@app.route('/social/status', methods=['GET'])
def social_status():
    return jsonify({'success': True, 'platforms': _social_public_status()})


# Navigates the browser to the platform's own login page. The random state
# tied to the session is what the callback checks, so a link someone else
# crafts can't connect an account to this one.
@app.route('/social/connect/<platform>', methods=['GET'])
def social_connect(platform):
    provider = SOCIAL_PROVIDERS.get(platform)
    if not provider or 'auth_url' not in provider:
        return jsonify({'success': False, 'error': 'Unknown platform'}), 404
    if not provider['configured']():
        return jsonify({'success': False, 'error': f"{SOCIAL_LABELS[platform]} isn't set up on the server yet."}), 400
    state = secrets.token_urlsafe(24)
    session['social_oauth'] = {'platform': platform, 'state': state}
    return redirect(provider['auth_url'](state, _social_redirect_uri(platform)))


@app.route('/social/connect/practice', methods=['POST'])
def social_connect_practice():
    with social_lock:
        saved = _social_load()
        saved['practice'] = {'accountName': 'Practice mode - nothing is posted', 'connectedAt': now_local().isoformat()}
        _social_save(saved)
    return jsonify({'success': True})


@app.route('/social/callback/<platform>', methods=['GET'])
def social_callback(platform):
    def back(**params):
        return redirect('/?' + urllib.parse.urlencode(dict(params, platform=platform)))

    saved_state = session.pop('social_oauth', None)
    provider = SOCIAL_PROVIDERS.get(platform)
    if not provider or 'finish' not in provider:
        return back(social='error', message='Unknown platform')
    if not saved_state or saved_state.get('platform') != platform or not secrets.compare_digest(
            str(saved_state.get('state', '')), str(request.args.get('state', ''))):
        return back(social='error', message="That connection attempt didn't match this browser session - try Connect again.")
    denied = request.args.get('error_description') or request.args.get('error')
    if denied:
        return back(social='error', message=str(denied)[:200])
    code = request.args.get('code')
    if not code:
        return back(social='error', message='No authorization code came back.')
    try:
        account = provider['finish'](code, _social_redirect_uri(platform))
    except SocialError as e:
        print(f"Social connect error ({platform}): {e}")
        return back(social='error', message=str(e)[:200])
    except Exception as e:
        print(f"Social connect error ({platform}): {e}")
        return back(social='error', message="Something went wrong finishing the connection.")
    account['connectedAt'] = now_local().isoformat()
    with social_lock:
        saved = _social_load()
        saved[platform] = account
        _social_save(saved)
    return back(social='connected')


@app.route('/social/disconnect/<platform>', methods=['POST'])
def social_disconnect(platform):
    if platform not in SOCIAL_PROVIDERS:
        return jsonify({'success': False, 'error': 'Unknown platform'}), 404
    with social_lock:
        saved = _social_load()
        saved.pop(platform, None)
        _social_save(saved)
    return jsonify({'success': True})


@app.route('/social/page', methods=['POST'])
def social_choose_page():
    data = request.json or {}
    with social_lock:
        saved = _social_load()
        account = saved.get('facebook')
        if not account or not any(p['id'] == data.get('pageId') for p in account.get('pages', [])):
            return jsonify({'success': False, 'error': 'Unknown page'}), 400
        account['pageId'] = data['pageId']
        account['accountName'] = next(p['name'] for p in account['pages'] if p['id'] == data['pageId'])
        _social_save(saved)
    return jsonify({'success': True})


def _social_log(entry):
    try:
        with open(SOCIAL_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(json.dumps(entry) + '\n')
    except OSError as e:
        print(f"Could not write the publish log: {e}")


# The only way anything is published. It runs when Francis presses Publish on
# the review screen: it needs confirm=true, posts exactly the text he sent
# (which he can edit there), and records every attempt. publishId makes a
# double click or a retry safe - a platform that already succeeded for that
# id returns its earlier result instead of posting a second time.
@app.route('/social/publish', methods=['POST'])
def social_publish():
    data = request.json or {}
    if data.get('confirm') is not True:
        return jsonify({'success': False, 'error': 'Publishing needs explicit confirmation.'}), 400
    text = str(data.get('text') or '').strip()
    platforms = [p for p in (data.get('platforms') or []) if p in SOCIAL_PROVIDERS]
    publish_id = str(data.get('publishId') or '').strip()[:80]
    if not text:
        return jsonify({'success': False, 'error': 'There is nothing to post.'}), 400
    if not platforms:
        return jsonify({'success': False, 'error': 'Choose at least one account.'}), 400
    if not publish_id:
        return jsonify({'success': False, 'error': 'Missing publish id.'}), 400

    results = []
    for platform in dict.fromkeys(platforms):
        with social_lock:
            cached = _social_recent_publishes.get((publish_id, platform))
            account = _social_load().get(platform)
            if cached:
                results.append(cached)
                continue
            result = {'platform': platform, 'label': SOCIAL_LABELS[platform]}
            try:
                if not account:
                    raise SocialError(f"{SOCIAL_LABELS[platform]} isn't connected.")
                if len(text) > SOCIAL_MAX_LENGTH[platform]:
                    raise SocialError(f"That's {len(text)} characters; {SOCIAL_LABELS[platform]} allows {SOCIAL_MAX_LENGTH[platform]}.")
                outcome = SOCIAL_PROVIDERS[platform]['publish'](account, text)
                result.update(success=True, id=outcome.get('id'), url=outcome.get('url'), practice=bool(outcome.get('practice')))
            except SocialError as e:
                result.update(success=False, error=str(e))
            except Exception as e:
                print(f"Social publish error ({platform}): {e}")
                result.update(success=False, error='Something went wrong while posting.')
            if result['success']:
                _social_recent_publishes[(publish_id, platform)] = result
            _social_log({
                'ts': now_local().isoformat(), 'platform': platform, 'success': result['success'],
                'practice': result.get('practice', False), 'postId': result.get('id'), 'error': result.get('error'),
                'agent': data.get('agent'), 'taskId': data.get('taskId'), 'textPreview': text[:200]
            })
            results.append(result)
    return jsonify({'success': True, 'results': results})


# --- Task plans ---------------------------------------------------------
# Every task gets a short plan when it is created: what will get done, the
# steps the agent will take, and what the final output is. Francis sees it on
# the task before pressing Start; it is handed to the agent when the task
# runs, and the agent's finished work then goes to Francis for review before
# the task counts as complete. One small call per task.
TASK_PLAN_CAPABILITIES = (
    "The agent can research with web search, look things up in the firm's uploaded library, "
    "write in chat, and create files (Word, Excel, PowerPoint, PDF or HTML). It cannot send emails "
    "or messages, post anywhere, log in to any account, run code, or change anything outside this "
    "app. Plan only what the agent can actually do here; if the real-world follow-through is up to "
    "Francis, say so in the final output (for example, \"Francis then posts it\")."
)


TASK_PLAN_SASHA_NOTE = (
    " Exception for Sasha: she can draft the text of a social media post. Once Francis has reviewed and "
    "approved it, the app publishes it to his connected social accounts - so a social post's plan should "
    "end with the finished draft, and its final output should say it is published after his approval "
    "rather than that he posts it himself. She still can't make images or video."
)


def _first_json_object(text):
    decoder = json.JSONDecoder()
    pos = text.find('{')
    while pos != -1:
        try:
            value, _ = decoder.raw_decode(text[pos:])
            if isinstance(value, dict):
                return value
        except ValueError:
            pass
        pos = text.find('{', pos + 1)
    raise ValueError('The model did not return a plan.')


@app.route('/tasks/plan', methods=['POST'])
def tasks_plan():
    try:
        data = request.json or {}
        agent = data.get('agent')
        task_text = str(data.get('task') or '').strip()[:3000]
        task_name = str(data.get('name') or '').strip()[:200]
        if agent not in ALL_AGENTS or not task_text:
            return jsonify({'success': False, 'error': 'An agent and a task are required.'}), 400

        side = 'personal' if data.get('context') == 'personal' else 'firm'
        notes = str(load_kb_notes().get(side) or '').strip()[:1500]
        role = get_agent_role_summary(agent)

        file_names = [str(n)[:120] for n in (data.get('files') or []) if str(n).strip()][:10]
        project_line = ''
        steps = [x for x in (data.get('project_steps') or []) if isinstance(x, dict)]
        if data.get('project_name') and steps:
            listing = "; ".join(f"{i + 1}. {str(x.get('name') or '')[:80]} ({str(x.get('agent') or '')})" for i, x in enumerate(steps[:12]))
            project_line = (
                f"\nThis task is one step in the project \"{str(data['project_name'])[:120]}\", whose steps run in "
                f"order: {listing}. Plan only this task's own step, building on the earlier steps' output."
            )

        prompt = (
            f"Write a short work plan for {agent.capitalize()}, one of Francis's AI assistants"
            f"{' (' + role + ')' if role else ''}.\n\n"
            f"Task: {task_name + ' - ' if task_name else ''}{task_text}{project_line}\n\n"
            + (f"What Francis has told us about his {'life' if side == 'personal' else 'firm'} (use it only to tailor the plan - "
               f"never restate these facts inside the steps, and never invent details beyond them):\n{notes}\n\n" if notes else '')
            + (f"Files Francis provided for this task - the agent is given them to work from: {', '.join(file_names)}\n\n" if file_names else '')
            + TASK_PLAN_CAPABILITIES + (TASK_PLAN_SASHA_NOTE if agent == 'sasha' else '') + "\n\n"
            "Return ONLY a JSON object with these keys:\n"
            "- goal: one or two sentences on what will get done and why it is worth doing.\n"
            "- steps: 3 to 6 steps in the order the agent will take them. Each step is an object with two keys: "
            "\"title\" (what the step is, as a short action of 2 to 6 words starting with a verb, no colon - for example "
            "\"Review the existing file\") and \"detail\" (one plain sentence on what will take place in that step).\n"
            "- output: one or two sentences on exactly what Francis receives at the end - the deliverable and its "
            "format (for example, a two-page Word document, or a short written summary in chat)."
        )
        response = claude_create(
            log_agent=agent, log_purpose='task_plan', model=CLAUDE_MODEL, max_tokens=1400,
            messages=[{'role': 'user', 'content': prompt}]
        )
        raw = _first_json_object("".join(b.text for b in response.content if getattr(b, 'type', None) == 'text'))
        goal = str(raw.get('goal') or '').strip()
        plan_steps = []
        for item in (raw.get('steps') or [])[:7]:
            if isinstance(item, dict):
                step_title = str(item.get('title') or '').replace(':', ' ').strip()
                step_detail = str(item.get('detail') or '').strip()
                step = f"{step_title}: {step_detail}" if step_title and step_detail else (step_title or step_detail)
            else:
                step = str(item).strip()
            if step:
                plan_steps.append(step)
        output = str(raw.get('output') or '').strip()
        if not (goal and plan_steps and output):
            raise ValueError('The plan came back incomplete.')
        return jsonify({'success': True, 'plan': {'goal': goal, 'steps': plan_steps, 'output': output}})
    except Exception as e:
        print(f"Task plan error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/classify-task-needed', methods=['POST'])
def classify_task_needed():
    try:
        data = request.json or {}
        message = str(data.get('message') or '').strip()
        if not message:
            return jsonify({'success': True, 'needs_task': False})
        attachment_names = [str(n)[:120] for n in (data.get('attachment_names') or []) if str(n).strip()][:10]

        response = claude_create(
            model=CLAUDE_MODEL,
            max_tokens=200,
            system=(
                "You triage incoming requests to a team of AI assistants. Decide whether the request "
                "needs a background task - deep research (several web searches) or generating a "
                "downloadable file - versus something answerable directly in a normal quick chat reply. "
                "You MUST answer by calling the classify tool exactly once, with no other output."
                + (" The user also attached files. Files shared only so they can be read, discussed, or "
                   "mined for an answer ('what does this say', 'what's the total', 'FYI, for reference') do "
                   "NOT need a task. If the user wants the assistants to DO something with the files - turn "
                   "them into a report or document, edit, convert, fill in, reconcile, or build something "
                   "from them, or any multi-step work - that needs a task." if attachment_names else '')
            ),
            messages=[{"role": "user", "content": message + (f"\n\n[Attached files: {', '.join(attachment_names)}]" if attachment_names else '')}],
            tools=[CLASSIFY_TASK_TOOL]
        )

        for block in response.content:
            if getattr(block, 'type', None) == 'tool_use' and block.name == 'classify':
                block_input = block.input or {}
                return jsonify({
                    'success': True,
                    'needs_task': bool(block_input.get('needs_task')),
                    'task_name': str(block_input.get('task_name') or '').strip()
                })

        return jsonify({'success': True, 'needs_task': False})
    except Exception as e:
        print(f"Classify task error: {e}")
        return jsonify({'success': False, 'needs_task': False, 'error': str(e)})


# Francis sends a file in chat while the agent has a task waiting on it (the agent asked for it,
# or it plainly belongs to that task): this decides which open task - if any - the files are for,
# so they go under that task instead of sitting loose. Forced tool choice, one short call.
MATCH_FILE_TASK_TOOL = {
    "name": "match",
    "description": "Say which task, if any, the files are for.",
    "input_schema": {
        "type": "object",
        "properties": {
            "task_id": {"type": "string", "description": "The id of the task the files belong to, or \"none\"."}
        },
        "required": ["task_id"]
    }
}


@app.route('/match-file-to-task', methods=['POST'])
def match_file_to_task():
    try:
        data = request.json or {}
        tasks = [t for t in (data.get('tasks') or []) if isinstance(t, dict) and t.get('id')][:12]
        names = [str(n)[:120] for n in (data.get('attachment_names') or []) if str(n).strip()][:10]
        if not tasks or not names:
            return jsonify({'success': True, 'task_id': None})
        recent = "\n".join(
            f"{'Francis' if m.get('role') == 'user' else 'Agent'}: {str(m.get('text') or '')[:700]}"
            for m in (data.get('recent') or [])[-7:] if isinstance(m, dict)
        )
        listing = "\n".join(f"- id {t['id']}: {str(t.get('name') or '')[:100]} - {str(t.get('task') or '')[:400]}" for t in tasks)
        prompt = (
            "Francis just sent files in chat. Decide whether they are for one of the agent's open tasks that haven't been "
            "started yet - for example the agent asked him for the file, was asking questions about it, or it's plainly the "
            "material the task needs. Files that are only for reading or reference with no task needing them are for none.\n\n"
            f"Recent conversation:\n{recent or '(none)'}\n\n"
            f"His message: {str(data.get('message') or '')[:600]}\nFiles: {', '.join(names)}\n\n"
            f"Open tasks:\n{listing}\n\nCall match with the id of the task the files are for, or \"none\"."
        )
        response = claude_create(
            log_agent=None, log_purpose='match_file_task', model=CLAUDE_MODEL, max_tokens=100,
            messages=[{'role': 'user', 'content': prompt}],
            tools=[MATCH_FILE_TASK_TOOL], tool_choice={'type': 'tool', 'name': 'match'}
        )
        block = next((b for b in response.content if getattr(b, 'type', None) == 'tool_use'), None)
        task_id = str((block.input or {}).get('task_id') or '').strip() if block else ''
        return jsonify({'success': True, 'task_id': task_id if task_id in {str(t['id']) for t in tasks} else None})
    except Exception as e:
        print(f"Match file to task error: {e}")
        return jsonify({'success': False, 'task_id': None, 'error': str(e)})


@app.route('/chat', methods=['POST'])
def chat():
    try:
        data = request.json
        agent = data.get('agent')
        message = data.get('message') or ''
        history = data.get('history', [])
        attachments = data.get('attachments', [])
        task_context = str(data.get('task_context', '') or '').strip()
        project_context = str(data.get('project_context', '') or '').strip()
        is_task_run = bool(data.get('is_task_run'))
        open_tasks = [t for t in (data.get('open_tasks') or []) if isinstance(t, dict) and t.get('id')][:10] if not data.get('is_task_run') else []
        # A Lana task that is preparing one interactive lesson (see create_lesson).
        is_lesson_run = bool(data.get('is_lesson')) and agent == 'lana' and is_task_run

        print(f"Agent: {agent}, Message: {message}")

        if not agent or (not message.strip() and not attachments):
            return jsonify({
                'success': False,
                'error': 'Missing agent or message'
            }), 400

        # Load agent system prompt
        system_static = get_agent_system_prompt(agent)
        system_notes = get_knowledge_base_context() + get_personal_knowledge_context()
        # Everything appended to system_prompt from here on changes message to
        # message, so it's sent after the cached blocks (see build_system_blocks).
        system_prompt = build_history_digest(history, agent)
        if agent == 'ashanti':
            system_prompt += get_calendar_context()
            system_prompt += (
                "\n\nWhen Francis asks you to add something to the calendar (including a quoted item whose "
                "speaker label ends in \"(add to calendar)\" - this covers both a To-Do and a Discussion Topic "
                "sent to you this way) and hasn't given a specific day or time, don't guess - "
                "reply with a short question asking when, naming the item by its own exact title from THIS "
                "quote (e.g. \"When do you want 'Follow up with vendor' on the calendar?\") - never a different "
                "item's title from earlier in the conversation, even one you're still waiting on an answer "
                "for. Each \"(add to calendar)\" quote is its own separate item with its own separate timing "
                "question - if an earlier one is still unresolved when a new one comes in, ask about the NEW "
                "one on its own terms (don't blend the two into one message, don't guess that the new answer "
                "applies to the old item or vice versa); circle back to the earlier one separately if needed. "
                "You MUST, in that exact same turn, also call suggest_quick_replies with exactly these "
                "options: \"Today\", \"Tomorrow\", \"This week\", \"Next week\", \"Specific date\" - this is "
                "not optional, do not send that question as plain text alone even once. Never call "
                "suggest_quick_replies without that accompanying question text. Once he answers, pick (or ask for) an actual date/time and "
                "call manage_calendar to create it - don't leave it unscheduled once you have enough to act on. "
                "Whenever Francis gives you a day but not a specific time (e.g. \"Today\", \"Tomorrow\", "
                "\"next Tuesday\"), pick the time yourself using the FREE TIME block in that same CALENDAR "
                "context - it already lists the actual open windows for each of the next 7 days, with business "
                "hours, every existing event, and (for today) the current time all already accounted for. Pick "
                "a start time from within one of those listed windows, never anywhere else - do NOT reason "
                "about business hours or conflicts yourself from the raw event list, and do NOT default to a "
                "flat time like 10 AM without checking it's actually inside a free window first; a wrong guess "
                "here is exactly how double-booking and running past 5 PM happen. The event's full duration "
                "(start through start+estimated time) must fit within a single free window - if the only "
                "windows left are shorter than the duration, that's not a valid slot even if it starts free. "
                "If the day Francis named has no free window at all long enough for this (FREE TIME will say "
                "\"fully booked\" or list only shorter gaps), do NOT call manage_calendar and do NOT silently "
                "double-book or run it past business hours - tell him that day's full (or too tight) and ask "
                "whether he'd like a different day/time instead, the same way you'd ask about any open "
                "question. A quoted \"(add to "
                "calendar)\" item always states its own estimated time (e.g. \"Estimated time: 30 minutes\") - "
                "use that exact duration for the event's end (start + that duration), never a guessed or "
                "default length. When you "
                "call manage_calendar to create it, set origin=\"todo\" if the quote's speaker label was "
                "exactly \"To-Do (add to calendar)\", or origin=\"discussion\" if it was some other name "
                "followed by \"(add to calendar)\" (a Discussion Topic's category) - this is what puts the "
                "right icon on it on the Calendar page, so get it right. Leave origin out entirely for any "
                "other event you create with no such quote. "
                "Everything on the calendar is a to-do. When Francis needs more time on something already on "
                "the calendar (e.g. it's still open and you offer to block more time for it), call "
                "manage_calendar create with that item's todo_id (shown as \"[to-do id: ...]\" next to its "
                "event above) - never a new differently-titled event. The new block is another split of the "
                "same to-do and mirrors it exactly."
            )

        if open_tasks:
            lines = []
            for t in open_tasks:
                files = ", ".join(str(f)[:80] for f in (t.get('files') or [])[:6])
                lines.append(
                    f"- id {t['id']} | \"{str(t.get('name') or '')[:100]}\" | status: {str(t.get('status') or '')} "
                    f"({'editable with refine_task' if t.get('editable') else 'already underway - not editable'}) | "
                    f"task: {str(t.get('task') or '')[:600]}"
                    + (f" | plan goal: {str(t['planGoal'])[:300]}" if t.get('planGoal') else '')
                    + (f" | files given for it: {files}" if files else '')
                )
            system_prompt += "\n\nYOUR EXISTING TASKS:\n" + "\n".join(lines) + OPEN_TASKS_RULES
        if agent == 'lana':
            system_prompt += get_lana_context() + LANA_LESSON_RULES
            if is_lesson_run:
                system_prompt += LANA_TASK_RUN_RULES

        # task_context is only sent when Francis is discussing one specific
        # task from a PROJECT (see the "Discuss with Manny" flow on a project
        # task's Edit button) - it may not even be your own task, since
        # whoever's coordinating the project is the one who fields this, not
        # necessarily whoever it's assigned to. project_context (when
        # present) lists every task in that same project, so you can judge
        # whether a change here affects any of the others.
        if task_context:
            if project_context:
                system_prompt += (
                    f"\n\nFrancis wants to discuss this specific task from a project: "
                    f"\"{task_context}\". This may not be your own task - you're "
                    f"fielding it because you coordinate the project it belongs to."
                )
            else:
                system_prompt += (
                    f"\n\nFrancis is talking about this existing task of yours: \"{task_context}\". "
                    f"It already exists in the app - his message is about THIS task, so never propose or create a new task for it. "
                    f"A request to change it or add to it (\"also add...\", \"make it...\") is a settled change: apply it with "
                    f"update_task right away."
                )
            if project_context:
                system_prompt += (
                    f"\n\nHere is the full project, for context on whether a change to "
                    f"the task above affects any of the others in it:\n\n{project_context}"
                )
            system_prompt += (
                f"\n\nOnce Francis actually settles on a change to the task described "
                f"above, call update_task with its new, detailed description reflecting "
                f"what you agreed on - don't call it while still just discussing options. "
                f"Always pair that call with a short reply of your own confirming what you "
                f"changed it to - never call update_task as your only output, since that "
                f"reply is the only visible confirmation Francis gets that it actually "
                f"happened. update_task only restates THAT one task; if the change would "
                f"meaningfully affect another task in the project too, say so in your "
                f"reply so Francis knows, but you can't update the other task from here."
            )

        # is_task_run means this message came from Francis clicking Start/Resume
        # on a Workspace task (see startProjectTask), not a normal typed
        # message - the model otherwise has no way to know that, and without
        # this it can (and has) narrated a multi-step plan in prose as if
        # walking through building something, without ever actually calling
        # create_file or pause_task, which the frontend then wrongly reads as
        # a finished result.
        if is_task_run:
            system_prompt += (
                "\n\nThis message is Francis starting a background task from his Workspace - he's "
                "expecting you to actually complete the work in this conversation, not describe a plan "
                "for how you'd do it. If what he asked for calls for a downloadable file (a document, "
                "spreadsheet, presentation, PDF, or HTML page), you must actually call create_file with "
                "the real, finished content before you're done - narrating that you're about to build it, "
                "checking a step works, or drafting it in some intermediate format is not the same as "
                "producing it, since you have no way to actually run code here; create_file is the only "
                "way a file gets made. If several file formats were asked for (e.g. a PDF, a Word doc, "
                "AND a PowerPoint), call create_file once per format - every file you produce this way "
                "gets attached to the task, across as many turns as it takes, so it's fine to make one or "
                "two now and the rest later rather than trying to force them all into a single reply. If "
                "you genuinely can't finish everything in this turn - either from using up your available "
                "web searches partway through research, or because there's more to generate than "
                "comfortably fits in one reply - call pause_task with a short note on what's done and "
                "what's left (e.g. \"PDF and Word doc are done, PowerPoint is next\"), rather than ending "
                "the turn without having either fully finished or explicitly paused. Francis's message may "
                "include a plan for the task - follow its steps and deliver its stated final output. When you "
                "finish, say briefly what you produced and where to find it; Francis reviews the work before "
                "the task counts as complete, so don't call it final or ask him to mark it complete."
            )
            if agent == 'sasha':
                system_prompt += (
                    "\n\nWhen the finished work is a social media post, call draft_social_post with the final "
                    "post text. You can't publish anything: Francis reviews your draft and, once he approves, "
                    "publishes it himself to his connected accounts. Never claim you posted it, and keep any "
                    "alternatives or notes in your reply rather than in the post text."
                )

        # Tools: quick replies and referrals always available. propose_task lets
        # any agent turn a single settled piece of their own work into a real,
        # standalone task - propose_project is for when it's genuinely several
        # tasks that belong together as one piece of coordinated work (Manny is
        # the one who assembles a project spanning several colleagues - see
        # REFER_TEAMMATE_TOOL/PROPOSE_PROJECT_TOOL for the referral-to-Manny
        # flow that gets him there).
        tools = [
            {"type": "web_search_20260209", "name": "web_search", "max_uses": 10},
            QUICK_REPLIES_TOOL,
            REFER_TEAMMATE_TOOL,
            PROPOSE_TASK_TOOL,
            PAUSE_TASK_TOOL
        ]
        # A project is work passed between colleagues, which Manny sets up.
        if agent == 'manny':
            tools.append(PROPOSE_PROJECT_TOOL)
        # Files only come out of a task, so the task's page can hold them and
        # Francis can review them - never from a plain chat.
        if is_task_run and not is_lesson_run:
            tools.append(CREATE_FILE_TOOL)
        if agent == 'lana':
            tools.append(PROPOSE_LANGUAGE_PLAN_TOOL)
            tools.append(ADD_LESSONS_TOOL)
            tools.append(ADD_EXERCISES_TOOL)
            if is_lesson_run:
                tools.append(CREATE_LESSON_SET_TOOL)
        if agent == 'sasha' and is_task_run:
            tools.append(DRAFT_SOCIAL_POST_TOOL)
        if task_context:
            tools.append(UPDATE_TASK_TOOL)
        if agent == 'ashanti':
            tools.append(MANAGE_CALENDAR_TOOL)
        if library_has_content():
            tools.append(SEARCH_LIBRARY_TOOL)
        editable_task_ids = {str(t['id']) for t in open_tasks if t.get('editable')}
        if editable_task_ids:
            tools.append(REFINE_TASK_TOOL)

        # Build the conversation, then fold any attachments (images, PDFs, or
        # extracted text from Word/Excel/etc.) into the final turn's content as
        # extra blocks alongside the typed message.
        claude_messages = build_claude_messages(history, agent, message)
        attachment_blocks = build_attachment_content_blocks(attachments)
        if attachment_blocks:
            last_turn = claude_messages[-1]
            text_content = last_turn['content'] if isinstance(last_turn['content'], str) else ''
            content_blocks = []
            if text_content.strip():
                content_blocks.append({'type': 'text', 'text': text_content})
            content_blocks.extend(attachment_blocks)
            last_turn['content'] = content_blocks

        # Get response from Claude. Raised from 4096 -> 8192 -> 16000: a heavy
        # multi-search research turn's real total (search-result processing
        # plus the model's own generation) can run well past what looks like
        # a generous ceiling - confirmed directly (a real turn hit stop_reason
        # 'max_tokens' at usage.output_tokens=9972, already above the previous
        # 8192 cap, with no tool calls to show for it).
        chat_system = build_system_blocks(system_static, system_notes, system_prompt)

        def ask_claude():
            return claude_create(
                log_agent=agent,
                log_purpose='chat',
                model=CLAUDE_MODEL,
                max_tokens=16000,
                system=chat_system,
                messages=claude_messages,
                tools=tools
            )

        response = ask_claude()

        # search_library is the one tool whose result the model has to see
        # before it can answer, so (unlike the others, which only record
        # what the agent chose to do) it gets a real round trip: run the
        # search, hand the passages back, and let the model continue. Capped
        # so a confused search loop can't run away.
        for _ in range(3):
            library_calls = [b for b in response.content
                             if getattr(b, 'type', None) == 'tool_use' and getattr(b, 'name', None) == 'search_library']
            if response.stop_reason != 'tool_use' or not library_calls:
                break
            tool_results = []
            for block in response.content:
                if getattr(block, 'type', None) != 'tool_use':
                    continue
                if block.name == 'search_library':
                    try:
                        result_text = run_library_search(block.input or {})
                    except Exception as search_error:
                        print(f"Library search error: {search_error}")
                        result_text = "The library search failed - answer without it and say you couldn't check the files."
                else:
                    result_text = 'OK'
                tool_results.append({'type': 'tool_result', 'tool_use_id': block.id, 'content': result_text})
            claude_messages.append({'role': 'assistant', 'content': response.content})
            claude_messages.append({'role': 'user', 'content': tool_results})
            response = ask_claude()

        # Extract response text - concatenate every text block, since a web search
        # turn interleaves text with server-side tool-use/result blocks rather than
        # returning a single text block.
        response_text = "".join(
            block.text for block in response.content if getattr(block, 'type', None) == 'text'
        )

        # A heavy web-search turn, or a large generated file (a multi-section
        # HTML page, a long document), can spend the whole token budget before
        # leaving room for the actual reply text. Rather than silently
        # returning nothing, surface that so the user isn't left waiting with
        # no visible outcome.
        hit_max_tokens_empty = not response_text.strip() and response.stop_reason == 'max_tokens'
        if hit_max_tokens_empty:
            response_text = "That took more room to work through than expected and I ran out of space to answer - try asking again, maybe split into smaller steps."

        # If the agent just asked a short multiple-choice clarifying question, it may
        # have called the (purely cosmetic) quick-replies tool to suggest tappable
        # options instead of leaving the user to type a free-text answer. We never
        # execute this tool or continue the turn for it - just read its input.
        quick_replies = []
        refer = None
        propose_project = None
        propose_tasks = []
        add_lessons = []
        add_exercises = []
        task_refinements = []
        created_files = []
        social_draft = None
        task_update = None
        calendar_update = None
        # A hard token-limit cutoff during a Started task leaves nothing to
        # show and no tool call for the model to have signaled a real pause
        # with - it just ran out of room mid-thought. Treat that as a pause
        # anyway so the existing auto-continue picks it back up, rather than
        # the task falsely landing on 'completed' with an apology as its
        # only "result".
        task_paused = hit_max_tokens_empty and is_task_run
        for block in response.content:
            if getattr(block, 'type', None) != 'tool_use':
                continue
            block_name = getattr(block, 'name', None)
            if block_name == 'suggest_quick_replies':
                options = (block.input or {}).get('options', [])
                quick_replies = [str(o).strip() for o in options if str(o).strip()][:5]
            elif block_name == 'create_file':
                block_input = block.input or {}
                file_type = str(block_input.get('file_type', '')).strip().lower()
                filename = str(block_input.get('filename', '')).strip() or 'document'
                file_content = block_input.get('content', '')
                file_theme = str(block_input.get('theme', '')).strip().lower() or DEFAULT_THEME
                file_primary_color = str(block_input.get('primary_color', '')).strip() or None
                file_accent_color = str(block_input.get('accent_color', '')).strip() or None
                if file_type in FILE_TYPE_MIME and str(file_content).strip():
                    try:
                        file_bytes = generate_file_bytes(file_type, file_content, file_theme, file_primary_color, file_accent_color)
                        primary_rgb, accent_rgb = resolve_theme(file_theme, file_primary_color, file_accent_color)
                        created_files.append({
                            'name': f"{filename}.{file_type}",
                            'mimeType': FILE_TYPE_MIME[file_type],
                            'data': base64.b64encode(file_bytes).decode('ascii'),
                            # Everything below isn't needed to download the file - it lets the
                            # frontend's Workspace preview panel reconstruct a styled preview
                            # (real docx/xlsx/pptx binaries can't be rendered in-browser cheaply).
                            'fileType': file_type,
                            'content': file_content,
                            'theme': file_theme,
                            'primaryColor': '#{:02X}{:02X}{:02X}'.format(*primary_rgb),
                            'accentColor': '#{:02X}{:02X}{:02X}'.format(*accent_rgb)
                        })
                    except Exception as e:
                        print(f"File generation error: {e}")
            elif block_name == 'refer_to_teammate':
                block_input = block.input or {}
                refer_agent = block_input.get('agent')
                if refer_agent in ALL_AGENTS and refer_agent != agent:
                    refer = {
                        'agent': refer_agent,
                        'summary': str(block_input.get('summary', '')).strip()
                    }
            elif block_name == 'propose_project':
                block_input = block.input or {}
                raw_tasks = block_input.get('tasks', [])
                # Only Manny assembles a project spanning several colleagues -
                # anyone else proposing one is limited to their own task, even
                # if the model tried to list others too.
                if agent != 'manny':
                    raw_tasks = [t for t in raw_tasks if isinstance(t, dict) and t.get('agent') == agent]
                project_tasks = [
                    {'agent': t.get('agent'), 'task': str(t.get('task', '')).strip(), 'name': str(t.get('name', '')).strip()}
                    for t in raw_tasks
                    if isinstance(t, dict) and t.get('agent') in ALL_AGENTS and str(t.get('task', '')).strip()
                ]
                project_name = str(block_input.get('name', '')).strip()
                project_summary = str(block_input.get('summary', '')).strip()
                if project_tasks and project_name:
                    if len({t['agent'] for t in project_tasks}) >= 2:
                        propose_project = {
                            'name': project_name,
                            'summary': project_summary,
                            'tasks': project_tasks
                        }
                    else:
                        # A project is work handed from one colleague to the
                        # next. Several steps by the same person are just one
                        # task (the plan lists the steps), so fold it into that.
                        steps = "\n".join(f"{i + 1}. {t['task']}" for i, t in enumerate(project_tasks))
                        merged_agent = project_tasks[0]['agent']
                        propose_tasks = [t for t in propose_tasks if t['agent'] != merged_agent]
                        propose_tasks.append({
                            'agent': merged_agent,
                            'task': (project_summary + "\n\n" if project_summary else '') + steps,
                            'name': project_name
                        })
            elif block_name == 'propose_language_plan' and agent == 'lana':
                block_input = block.input or {}
                modules = []
                for m in (block_input.get('modules') or [])[:8]:
                    if not isinstance(m, dict) or not str(m.get('name') or '').strip():
                        continue
                    module_lessons = []
                    for x in (m.get('lessons') or [])[:12]:
                        if isinstance(x, dict) and str(x.get('task') or '').strip() and str(x.get('name') or '').strip():
                            module_lessons.append({
                                'name': str(x['name']).strip()[:100], 'task': str(x['task']).strip(),
                                'topics': [str(t).strip()[:120] for t in (x.get('topics') or []) if str(t).strip()][:8],
                                'review': bool(x.get('review'))
                            })
                    modules.append({
                        'name': str(m['name']).strip()[:100], 'level': str(m.get('level') or '').strip()[:20], 'days': str(m.get('days') or '').strip()[:40],
                        'goal': str(m.get('goal') or '').strip()[:300], 'lessons': module_lessons,
                        'roadmap': [str(r).strip()[:200] for r in (m.get('roadmap') or []) if str(r).strip()][:8]
                    })
                plan_name = str(block_input.get('name') or '').strip()[:120]
                if modules and modules[0]['lessons'] and plan_name:
                    propose_project = {
                        'name': plan_name, 'summary': str(block_input.get('summary') or '').strip(),
                        # Only the first module becomes tasks now (plus its quiz, added by the app);
                        # a roadmap module becomes lessons later, when Lana builds it.
                        'tasks': [dict(l, agent='lana') for l in modules[0]['lessons']],
                        'outline': {'modules': modules}, 'kind': 'language',
                        'language': str(block_input.get('language') or '').strip()[:40],
                        'languageCode': str(block_input.get('language_code') or '').strip()[:12]
                    }
            elif block_name == 'create_lesson' and is_lesson_run:
                lesson = _clean_lesson(block.input or {})
                if lesson:
                    lesson_json = json.dumps(lesson, ensure_ascii=False)
                    created_files.append({
                        'name': lesson['title'], 'mimeType': 'application/json',
                        'data': base64.b64encode(lesson_json.encode('utf-8')).decode('ascii'),
                        'fileType': 'lesson', 'content': lesson_json
                    })
            elif block_name == 'refine_task':
                block_input = block.input or {}
                refine_id = str(block_input.get('task_id') or '').strip()
                refine_text = str(block_input.get('task') or '').strip()
                if refine_id in editable_task_ids and refine_text:
                    task_refinements = [r for r in task_refinements if r['task_id'] != refine_id]
                    task_refinements.append({'task_id': refine_id, 'task': refine_text, 'name': str(block_input.get('name') or '').strip()[:100]})
            elif block_name == 'create_lesson_set' and is_lesson_run:
                lesson_files = _build_lesson_set(block.input or {})
                if lesson_files:
                    created_files.extend(lesson_files)
            elif block_name == 'add_exercises' and agent == 'lana':
                for x in ((block.input or {}).get('exercises') or [])[:3]:
                    if isinstance(x, dict) and str(x.get('title') or '').strip() and str(x.get('focus') or '').strip():
                        add_exercises.append({
                            'title': str(x['title']).strip()[:100], 'focus': str(x['focus']).strip()[:400],
                            'style': str(x.get('style') or 'mixed').strip()[:20]
                        })
            elif block_name == 'add_lessons' and agent == 'lana':
                for x in ((block.input or {}).get('lessons') or [])[:4]:
                    if isinstance(x, dict) and str(x.get('task') or '').strip() and str(x.get('name') or '').strip():
                        add_lessons.append({
                            'task': str(x['task']).strip(), 'name': str(x['name']).strip()[:100],
                            'topics': [str(t).strip()[:120] for t in (x.get('topics') or []) if str(t).strip()][:8],
                            'module': str(x.get('module') or '').strip()[:100]
                        })
            elif block_name == 'propose_task':
                block_input = block.input or {}
                task_text = str(block_input.get('task', '')).strip()
                task_name = str(block_input.get('name', '')).strip()
                # Only Manny hands tasks to other colleagues; everyone else's
                # task is their own, whatever the model put in `agent`.
                assignee = block_input.get('agent')
                if agent != 'manny' or assignee not in ALL_AGENTS:
                    assignee = agent
                if task_text and task_name:
                    # A repeat call for the same assignee refines that proposal.
                    propose_tasks = [t for t in propose_tasks if t['agent'] != assignee]
                    propose_tasks.append({'agent': assignee, 'task': task_text, 'name': task_name})
            elif block_name == 'draft_social_post' and agent == 'sasha' and is_task_run:
                block_input = block.input or {}
                draft_text = str(block_input.get('text', '')).strip()[:5000]
                if draft_text:
                    social_draft = {
                        'text': draft_text,
                        'platforms': [x for x in (block_input.get('platforms') or []) if x in ('linkedin', 'facebook')]
                    }
            elif block_name == 'update_task' and task_context:
                block_input = block.input or {}
                new_task_text = str(block_input.get('task', '')).strip()
                new_task_name = str(block_input.get('name', '')).strip()
                if new_task_text:
                    task_update = {'task': new_task_text, 'name': new_task_name}
            elif block_name == 'pause_task':
                block_input = block.input or {}
                reason = str(block_input.get('reason', '')).strip()
                if reason:
                    task_paused = True
                    # The reason IS the reply in this case - there's rarely
                    # separate text alongside this tool call, so this is the
                    # same fallback pattern as the max_tokens case above.
                    if not response_text.strip():
                        response_text = reason
            elif block_name == 'manage_calendar' and agent == 'ashanti':
                block_input = block.input or {}
                action = block_input.get('action')
                try:
                    with todos_lock, calendar_lock:
                        if action == 'create':
                            title = str(block_input.get('title', '')).strip()
                            start = str(block_input.get('start', '')).strip()
                            todos_all = load_todos()
                            todo_id_in = str(block_input.get('todo_id', '') or '').strip()
                            todo = next((t for t in todos_all if t['id'] == todo_id_in), None) if todo_id_in else None
                            if start and (todo or title):
                                now_iso = now_local().isoformat()
                                origin = block_input.get('origin')
                                end = str(block_input.get('end', '') or '').strip() or start
                                events = load_calendar_events()
                                if todo is None:
                                    # Nothing goes on the calendar that isn't a
                                    # to-do - a brand-new item gets its own.
                                    try:
                                        minutes = int((datetime.fromisoformat(end) - datetime.fromisoformat(start)).total_seconds() // 60)
                                    except (ValueError, TypeError):
                                        minutes = 0
                                    todo = {
                                        'id': uuid.uuid4().hex, 'userId': current_user_id(), 'title': title,
                                        'details': str(block_input.get('description', '') or '').strip() or title,
                                        'estimatedMinutes': minutes if minutes > 0 else 30,
                                        'completed': False, 'createdAt': now_iso, 'completedAt': None,
                                        'calendarEventId': None, 'calendarEventIds': [], 'attachments': [],
                                        'priority': 'medium', 'personal': False, 'notFinished': False
                                    }
                                    todos_all.append(todo)
                                new_event = {
                                    'id': uuid.uuid4().hex,
                                    'userId': current_user_id(),
                                    'title': todo['title'],
                                    'description': todo.get('details') or '',
                                    'location': str(block_input.get('location', '') or '').strip(),
                                    'start': start,
                                    'end': end,
                                    'allDay': bool(block_input.get('all_day', False)),
                                    'source': 'internal',
                                    'origin': origin if origin in ('todo', 'discussion') else 'todo',
                                    'externalUid': None,
                                    'status': 'confirmed',
                                    'createdAt': now_iso,
                                    'updatedAt': now_iso,
                                    'attachments': todo.get('attachments') or []
                                }
                                # Already on the calendar? Then this block is
                                # another split of the same to-do, tied to the
                                # block it's currently on.
                                ids = _normalize_todo_event_ids(todo)
                                current = next((e for e in events if e['id'] == todo.get('calendarEventId')), None)
                                if todo.get('notFinished'):
                                    current = None  # its old card was missed - this is a fresh scheduling, not a split
                                if current:
                                    new_event['splitOf'] = current.get('splitOf') or current['id']
                                events.append(new_event)
                                if current:
                                    _renumber_split_group(events, new_event)
                                    new_event = next(e for e in events if e['id'] == new_event['id'])
                                else:
                                    todo['calendarEventId'] = new_event['id']
                                if new_event['id'] not in ids:
                                    ids.append(new_event['id'])
                                todo['notFinished'] = False
                                save_calendar_events(events)
                                save_todos(todos_all)
                                calendar_update = {'action': 'create', 'event': new_event, 'todo_id': todo['id']}
                        elif action == 'update':
                            events = load_calendar_events()
                            ev = next((e for e in events if e['id'] == block_input.get('event_id')), None)
                            if ev:
                                try:
                                    orig_start = datetime.fromisoformat(ev['start']) if ev.get('start') else None
                                    orig_end = datetime.fromisoformat(ev['end']) if ev.get('end') else None
                                    orig_duration = (orig_end - orig_start) if (orig_start and orig_end) else None
                                except (ValueError, TypeError):
                                    orig_duration = None
                                for field in ('title', 'description', 'location', 'start', 'end'):
                                    if block_input.get(field):
                                        ev[field] = str(block_input[field]).strip()
                                if 'all_day' in block_input:
                                    ev['allDay'] = bool(block_input['all_day'])
                                # Belt-and-suspenders against a model-supplied
                                # start/end landing end-before-start (or some
                                # other nonsense) - falls back to the event's
                                # own duration from before this update rather
                                # than writing an invalid range.
                                if not ev.get('allDay'):
                                    try:
                                        new_start = datetime.fromisoformat(ev['start'])
                                        new_end = datetime.fromisoformat(ev['end']) if ev.get('end') else None
                                        if not new_end or new_end <= new_start:
                                            ev['end'] = (new_start + (orig_duration or timedelta(minutes=30))).isoformat()
                                    except (ValueError, TypeError):
                                        pass
                                ev['updatedAt'] = now_local().isoformat()
                                # Text changes belong to the to-do and reach
                                # every split of it; times only move this block.
                                linked = _todo_for_event(load_todos(), ev)
                                if linked and (block_input.get('title') or block_input.get('description')):
                                    todos_all = load_todos()
                                    linked = _todo_for_event(todos_all, ev)
                                    if block_input.get('title'):
                                        linked['title'] = ev['title']
                                    if block_input.get('description'):
                                        linked['details'] = ev['description']
                                    _sync_todo_to_events(linked, events)
                                    save_todos(todos_all)
                                save_calendar_events(events)
                                calendar_update = {'action': 'update', 'event': ev}
                        elif action == 'move_to_next_available':
                            events = load_calendar_events()
                            ev = next((e for e in events if e['id'] == block_input.get('event_id')), None)
                            if ev and not ev.get('allDay'):
                                try:
                                    orig_start = datetime.fromisoformat(ev['start'])
                                    orig_end = datetime.fromisoformat(ev['end']) if ev.get('end') else orig_start + timedelta(minutes=30)
                                    duration = orig_end - orig_start
                                except (ValueError, KeyError, TypeError):
                                    duration = timedelta(minutes=30)
                                after_raw = str(block_input.get('after', '')).strip()
                                try:
                                    search_from = datetime.fromisoformat(after_raw) if after_raw else now_local()
                                except ValueError:
                                    search_from = now_local()
                                slot = _find_next_available_event_slot(events, ev['id'], search_from, duration)
                                if slot:
                                    slot_start, slot_end = slot
                                    ev['start'] = slot_start.isoformat()
                                    ev['end'] = slot_end.isoformat()
                                    ev['updatedAt'] = now_local().isoformat()
                                    save_calendar_events(events)
                                    calendar_update = {'action': 'update', 'event': ev}
                        elif action == 'delete':
                            events = load_calendar_events()
                            ev = next((e for e in events if e['id'] == block_input.get('event_id')), None)
                            if ev:
                                todos_all = load_todos()
                                linked = _todo_for_event(todos_all, ev)
                                children = [e for e in events if e.get('splitOf') == ev['id']]
                                events = [e for e in events if e['id'] != ev['id']]
                                if children:
                                    # The head of a split went - the earliest
                                    # remaining block takes over as head.
                                    children.sort(key=lambda e: e.get('start') or '')
                                    new_head = children[0]
                                    new_head.pop('splitOf', None)
                                    for c in children[1:]:
                                        c['splitOf'] = new_head['id']
                                    if linked:
                                        ids = _normalize_todo_event_ids(linked)
                                        if new_head['id'] not in ids:
                                            ids.append(new_head['id'])
                                    _renumber_split_group(events, new_head)
                                elif ev.get('splitOf'):
                                    sibling = next((e for e in events if e['id'] == ev['splitOf'] or e.get('splitOf') == ev['splitOf']), None)
                                    if sibling:
                                        _renumber_split_group(events, sibling)
                                if linked:
                                    ids = [i for i in _normalize_todo_event_ids(linked) if i != ev['id']]
                                    linked['calendarEventIds'] = ids
                                    if linked.get('calendarEventId') == ev['id']:
                                        linked['calendarEventId'] = ids[-1] if ids else None
                                    save_todos(todos_all)
                                save_calendar_events(events)
                                calendar_update = {'action': 'delete', 'event_id': ev['id']}
                except Exception as e:
                    print(f"Calendar tool error: {e}")

        return jsonify({
            'success': True,
            'response': response_text,
            'quick_replies': quick_replies,
            'refer': refer,
            'propose_project': propose_project,
            'propose_tasks': propose_tasks,
            'add_lessons': add_lessons,
            'add_exercises': add_exercises,
            'task_refinements': task_refinements,
            'created_files': created_files,
            'social_draft': social_draft,
            'task_update': task_update,
            'calendar_update': calendar_update,
            'task_paused': task_paused
        })
    
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

# Purely cosmetic tool - never executed, never given a tool_result, and the turn
# never continues for it. It exists only so the model can hand the UI a small set
# of tappable options when it asks a short, closed-set clarifying question (e.g.
# "which sport?"), instead of the user having to type a free-text reply.
SEARCH_LIBRARY_TOOL = {
    "name": "search_library",
    "description": (
        "Look things up in the business's connected files (spreadsheets, documents) - the library. Use it whenever "
        "an answer depends on specifics from those files: exact figures, dates, a client's numbers, a particular "
        "row. Never guess or estimate such details from memory. Query: the distinctive words, names or dates as "
        "they'd appear in the file (e.g. a client name, or a date like 2026-10-03). For 'latest', 'most recent', "
        "'yesterday' or 'current' questions about a file that grows over time, set where='latest' to get its most "
        "recent entries. where='start' returns the beginning of a file. If a result says a file has more parts, "
        "pass part=N to read a specific part. You may search more than once to refine."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "query": {"type": "string", "description": "Words, names or dates to find. Optional when where or part is given."},
            "where": {"type": "string", "enum": ["match", "latest", "start"], "description": "match (default): best-matching passages. latest: the file's most recent (last) entries. start: the beginning."},
            "file": {"type": "string", "description": "Optional part of a file name to limit the search to one file."},
            "part": {"type": "integer", "description": "Read this specific numbered part of the (first matching) file."}
        }
    }
}

QUICK_REPLIES_TOOL = {
    "name": "suggest_quick_replies",
    "description": (
        "Call this ONLY when your reply ends on a short clarifying question that has "
        "a small, natural set of concrete answers (e.g. asking which of a few named "
        "things the user means). Provide 2-5 short options as the user might tap "
        "them, each just a few words - not full sentences. Do not call this for "
        "open-ended questions, or when there isn't a genuinely short list of likely "
        "answers."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "options": {
                "type": "array",
                "items": {"type": "string"},
                "minItems": 2,
                "maxItems": 5
            }
        },
        "required": ["options"]
    }
}

# Acted on by the frontend: shows a "Speak to [Name]" button that hands the
# target agent a context summary and takes Francis straight to their chat -
# either a single-specialist referral, or a referral to Manny when the work
# spans several colleagues (see the tool description for both cases).
REFER_TEAMMATE_TOOL = {
    "name": "refer_to_teammate",
    "description": (
        "Call this in either of two cases. (1) What Francis needs is squarely one specific "
        "colleague's domain - not a personal/hobby topic (that's the redirect behavior above), "
        "an actual task - and you have nothing further to contribute, so HE should go talk to "
        "them next. (2) The work genuinely needs several different colleagues' parts "
        "coordinated into a real plan, not just you or one other specialist - refer him to "
        "MANNY specifically (not each colleague individually), since working out who's needed "
        "and setting up each one's task (a project when the work passes from one colleague to the "
        "next in order) is his job, not something to assemble piecemeal from your own chat. If the "
        "work is entirely yours to do, however many steps it takes, don't refer it away - call "
        "propose_task instead. In every case, this shows Francis a "
        "\"Speak to [Name]\" button; if he clicks it, he's taken straight to that colleague's own "
        "1:1 chat with your summary already given to them as context, so he never repeats "
        "himself and they pick up right where you left off. Always pair this call with a short "
        "reply of your own explaining the referral - never call it as your only output."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "agent": {"type": "string", "enum": ALL_AGENTS},
            "summary": {
                "type": "string",
                "description": "2-4 sentences, written for that colleague to read as their starting context - what Francis needs and the relevant parts of this conversation, so they can continue without Francis repeating himself."
            }
        },
        "required": ["agent", "summary"]
    }
}

# Ashanti-only. The CALENDAR context block in her system prompt lists each
# upcoming event's id - required for update/delete, unused for create.
# Actually executed server-side in /chat (unlike the cosmetic quick-replies
# tool), so a real event exists the moment she calls this, no separate
# accept step.
MANAGE_CALENDAR_TOOL = {
    "name": "manage_calendar",
    "description": (
        "Create, update, move, or delete an event on Francis's Offload calendar. Only call this once Francis has "
        "actually confirmed the event or change - not while still discussing options or times. Use "
        "action=\"create\" to put a block on the calendar (needs start, plus todo_id for an existing to-do - e.g. extra time for one already on the calendar - or a title for a brand-new item), \"update\" to change an existing one to a "
        "SPECIFIC time Francis actually named (needs event_id plus whichever fields changed), "
        "\"move_to_next_available\" when Francis instead just wants it moved to the next open opening "
        "(optionally \"sometime next week\" or similar - he isn't naming an exact time), or \"delete\" to "
        "remove one (needs event_id). For move_to_next_available, do NOT compute the start/end time yourself - "
        "you have no way to know what's actually free, and guessing produces exactly the kind of broken event "
        "(end before start, wrong duration) this action exists to avoid. Just resolve whatever Francis said "
        "about timing (\"next week\", \"after Friday\", or nothing at all) into the `after` field and let the "
        "server find the real opening."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "action": {"type": "string", "enum": ["create", "update", "move_to_next_available", "delete"]},
            "event_id": {
                "type": "string",
                "description": "Required for update/move_to_next_available/delete - the event's id from the CALENDAR context above."
            },
            "todo_id": {
                "type": "string",
                "description": (
                    "action=\"create\" only. Everything on the calendar is a to-do. When the block you're adding "
                    "is for a to-do that already exists - including more time for one already on the calendar "
                    "(its \"[to-do id: ...]\" is shown next to its event in the CALENDAR context) - pass that id. "
                    "The new block becomes another split of that same to-do: it mirrors the to-do's title and "
                    "details, opening any block shows all the splits, edits reach all of them, and checking one "
                    "off completes the to-do. Never make a separate new item for time on an existing to-do. "
                    "Only omit this (and give a title) for something that truly is a new item; it gets its own to-do."
                )
            },
            "title": {"type": "string", "description": "action=\"create\" without todo_id, or action=\"update\"."},
            "start": {
                "type": "string",
                "description": "action=\"create\" and action=\"update\". ISO datetime like \"2026-09-22T14:00:00\", or just a date \"2026-09-22\" for an all-day event."
            },
            "end": {
                "type": "string",
                "description": "action=\"create\" and action=\"update\". Same format as start. Defaults to start if omitted."
            },
            "after": {
                "type": "string",
                "description": (
                    "action=\"move_to_next_available\" only. ISO date/datetime to start searching from - "
                    "e.g. next Monday's date for \"move it to next week\", or omit entirely for \"as soon as "
                    "possible\"/no timing mentioned. The server finds the actual first open business-hours slot "
                    "at or after this point long enough for the event's own existing duration - never pass a "
                    "guessed end time or duration here."
                )
            },
            "all_day": {"type": "boolean"},
            "location": {"type": "string"},
            "description": {"type": "string"},
            "origin": {
                "type": "string",
                "enum": ["todo", "discussion"],
                "description": (
                    "Only for action=\"create\", and only when this event comes from a quoted item whose "
                    "speaker label ends in \"(add to calendar)\": \"todo\" if that label was \"To-Do (add to "
                    "calendar)\", \"discussion\" if it was some other name (a Discussion Topic's category) "
                    "followed by \"(add to calendar)\". Omit entirely for anything else - a plain request with "
                    "no such quote."
                )
            }
        },
        "required": ["action"]
    }
}

# Acted on by the frontend. Posts a proposal card with an Accept Task button
# instead of a plain message. Accepting sends the task straight into the
# agent's own Workspace as a standalone item - no project wrapper, since it
# isn't grouped with anything else. This is how a single settled piece of
# work turns into actual assigned work, the same "Accept card IS the sign-off"
# checkpoint propose_project uses, just without the project overhead for the
# common case of just one thing.
PROPOSE_TASK_TOOL = {
    "name": "propose_task",
    "description": (
        "Call this once a real piece of work Francis wants has been fully thought through - you "
        "know exactly what needs to happen, with nothing meaningful left to figure out. A task is "
        "one colleague's piece of work, however many steps it takes them - the steps go in their "
        "own plan, so several steps by the same person are still just one task. You don't need "
        "Francis to separately say \"lock it in\" first, proposing it (with its own Accept card) "
        "IS how he signs off. If the work is someone else's, or needs several colleagues, don't "
        "call this - use refer_to_teammate to send Francis to them (to Manny if it spans several). "
        "Exception, Manny only: when Francis brings you work for several colleagues that does NOT "
        "have to happen in a particular order, call this once per colleague, setting `agent` to who "
        "does it - each becomes that colleague's own task. If each colleague's piece builds on the "
        "previous one's output, that's a project (propose_project), not several tasks. Do NOT call "
        "this while still being worked out - only once it's actually settled. Do NOT call it again "
        "for every small clarifying detail Francis asks about after already proposing it - just "
        "answer the question directly. If you call this again for the same person after an earlier "
        "proposal that Francis hasn't accepted yet, it UPDATES that proposal in place (it does not "
        "create a second one). If one of your EXISTING TASKS (listed in your context) already covers this - Francis is "
        "discussing it, or answering your questions about it - do NOT propose a new task: use refine_task on that task."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "task": {
                "type": "string",
                "description": "One clear, specific description of the task, written so whoever does it can pick it up and act on it in their own Workspace without needing to re-read this whole conversation."
            },
            "name": {
                "type": "string",
                "description": "A short 3-6 word title for this task, e.g. \"Draft Q3 Newsletter\" - not a sentence, just enough to identify it at a glance in a list."
            },
            "agent": {
                "type": "string",
                "enum": ALL_AGENTS,
                "description": "Manny only: which colleague does this task. Everyone else leaves this out - the task is always their own."
            }
        },
        "required": ["task", "name"]
    }
}

# Acted on by the frontend. Posts a proposal card with an Accept Project
# button instead of a plain message. Only Manny is offered this tool, since a
# project is work handed from one colleague to the next. Accepting sends each
# listed colleague's task into their own Workspace, in the order listed - and
# a later task can't be started until the one before it is completed (then
# it's handed that step's output). The Accept/Pass card itself is the "before
# work begins" checkpoint.
PROPOSE_PROJECT_TOOL = {
    "name": "propose_project",
    "description": (
        "Call this ONLY for work that passes between different colleagues in order - each one's "
        "piece builds on what the previous colleague produced (for example: Mark researches leads "
        "and writes a marketing plan, THEN Kat drafts the creative from that plan, THEN Sasha builds "
        "the social post from Kat's draft). A project is that ordered chain of tasks across at least "
        "TWO different colleagues, listed in the exact order they must happen; the app will not let "
        "a later task start until the one before it is done, and hands it the earlier output. Do NOT "
        "use it for several steps by the same colleague - that's a single task with a plan. Do NOT use "
        "it when the colleagues' pieces are independent (they could be done at the same time or in any "
        "order) - propose one task per colleague with propose_task instead. Call it only once the plan "
        "is fully thought through, with nothing left undecided. Do NOT call it again for every small "
        "clarifying detail Francis asks about after a plan was already proposed - just answer the "
        "question directly. If you call this again after an earlier proposal in the same conversation "
        "that Francis hasn't accepted yet, it UPDATES that same proposal in place - so only do that "
        "when something in the plan actually changed, and carry forward every detail from the prior "
        "version that's still accurate."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "name": {
                "type": "string",
                "description": "A short project name, e.g. \"Spring Lead Campaign\"."
            },
            "summary": {
                "type": "string",
                "description": "1-3 sentences, in your own voice, summarizing the plan and how the work passes from one colleague to the next - this is the message shown to Francis alongside the task breakdown."
            },
            "tasks": {
                "type": "array",
                "description": "The tasks in the exact order they must happen, each for a colleague - each starts only after the previous one is completed.",
                "items": {
                    "type": "object",
                    "properties": {
                        "agent": {"type": "string", "enum": ALL_AGENTS},
                        "task": {
                            "type": "string",
                            "description": "One clear, specific task for this colleague, written so they can pick it up and act on it in their own Workspace without needing to re-read this whole conversation. Say what it builds on from the previous colleague's work."
                        },
                        "name": {
                            "type": "string",
                            "description": "A short 3-6 word title for this task, e.g. \"Draft Q3 Newsletter\" - not a sentence, just enough to identify it at a glance in a list."
                        }
                    },
                    "required": ["agent", "task", "name"]
                },
                "minItems": 2
            }
        },
        "required": ["name", "summary", "tasks"]
    }
}

# Acted on by the frontend, and only offered to Sasha while she is running a
# task. It does NOT publish anything: it hands the app the finished post text,
# which is shown on the task's review screen. Francis edits it there if he
# likes and presses Publish himself (see /social/publish) - no tool lets an
# agent post.
DRAFT_SOCIAL_POST_TOOL = {
    "name": "draft_social_post",
    "description": (
        "Call this when the final output of the task you are running is a social media post. Give "
        "the finished post text exactly as it should appear - ready to publish, with no commentary, "
        "options or explanations around it (those go in your reply). It does not post anything: "
        "Francis reviews and edits the draft, then publishes it himself from the task's review "
        "screen. Never say or imply that you posted it. Text only - you cannot attach images or video."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "text": {
                "type": "string",
                "description": "The complete post, exactly as it should be published."
            },
            "platforms": {
                "type": "array",
                "items": {"type": "string", "enum": ["linkedin", "facebook"]},
                "description": "Which networks this wording is written for. Leave out if it works for any."
            }
        },
        "required": ["text"]
    }
}

# Acted on by the frontend: only offered on a 1:1 turn where Francis is
# discussing one specific task from a PROJECT (the "Discuss with Manny" flow
# on a project task's Edit button - task_context/project_context carry the
# task's own text and the rest of its project so this can be offered every
# turn about it, not just the first). Updates that task's own text on its
# Workspace card in place, even though the chat this runs in may not be the
# task's own owning agent's chat - a project task is always discussed with
# whoever coordinates the project, not necessarily whoever it's assigned to.
UPDATE_TASK_TOOL = {
    "name": "update_task",
    "description": (
        "Call this only once Francis has actually settled on a change to the task "
        "described above - not while you're still discussing options or asking "
        "clarifying questions. Restates the task in its new, detailed form based on "
        "what you agreed on, written clearly and specifically enough for whoever's "
        "doing it to pick up and act on without re-reading this conversation. This "
        "replaces that one task's description - it does not create a new task, and "
        "it cannot update any OTHER task in the project (if the change affects "
        "another one too, say so in your reply instead)."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "task": {
                "type": "string",
                "description": "The task's full new description, reflecting the settled change."
            },
            "name": {
                "type": "string",
                "description": "A short 3-6 word title for the updated task, e.g. \"Draft Q3 Newsletter\" - not a sentence, just enough to identify it at a glance in a list."
            }
        },
        "required": ["task", "name"]
    }
}

REFINE_TASK_TOOL = {
    "name": "refine_task",
    "description": (
        "Update one of YOUR EXISTING tasks (listed in your context as editable) when what you and Francis have been "
        "discussing about that task has settled into a clearer or changed version of it - new details, answers to your "
        "questions, a narrowed or expanded scope. It rewrites that task's description in place (its plan is redrafted); it "
        "does NOT create a new task, and the files Francis sent for it stay attached. Use this instead of propose_task "
        "whenever the work is the same task. Call it only once the change has actually settled, and always say in your "
        "reply that you've updated the existing task."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "task_id": {"type": "string", "description": "The id of the existing task, exactly as listed."},
            "task": {"type": "string", "description": "The task's full new description, specific enough to act on without re-reading the conversation."},
            "name": {"type": "string", "description": "A short 3-6 word title."}
        },
        "required": ["task_id", "task", "name"]
    }
}

OPEN_TASKS_RULES = (
    "\n\nThese tasks already exist for you in the app (you or Francis created them earlier). If what Francis is "
    "saying - discussing, answering your questions, handing you the file you asked for - is about one of them, the "
    "conversation IS about that task: never propose a new task for it. Files he sends in that conversation belong to "
    "it (the app attaches them). When the discussion changes or sharpens the task, call refine_task for the one marked "
    "editable and tell him you've updated the existing task; if it's already running or in review, say so and work "
    "from it instead. Only propose a new task for genuinely separate work."
)


CONVERSATION_STYLE_INSTRUCTIONS = """

## Conversation Style (applies to every reply)

- Reply like a real person in a chat, not a bot reciting copy. Match the length and energy of what was just said.
- Short casual messages ("hey", "thanks", "lol") get a sentence or two at most - never a paragraph, a list, or a recap of your role.
- Small talk gets small talk back: if asked how you are, actually answer briefly, in character, and ask it back. Don't skip past it to work.
- Don't steer to business yourself. With only a bare "hey" or pure small talk, open the door ("What's on your mind?") - never a pointed intake question out of nowhere ("How's the practice running?").
- Once the user gives you something real to react to - busy, stressed, working on something, a problem, even offhand ("ugh, mondays") - engage with that specific thing like a coworker: ask what's on their plate, offer to take something off it, or offer to loop in the right teammate. You're colleagues, not a vendor waiting to be asked.
- Go long or ask unprompted questions only when the message calls for it, and even then work through substantive topics in short exchanges (see Iterative Conversation below), not one giant message.
- Don't reintroduce yourself or restate your title unless it's the first message or you're asked who you are. Don't open with a greeting once the conversation is underway.
- Use the conversation history to stay consistent and avoid repeating yourself or re-asking things.
- Notice how the user writes (length, formality, punctuation, slang) and drift toward it - short and lowercase gets loose, formal gets tight - the way coworkers naturally mirror each other.
- Whenever your reply ends on a short clarifying question with a small, concrete set of likely answers (which of a few named things they mean), you MUST also call `suggest_quick_replies` in that same turn with 2-5 short options. Skip it only for open-ended questions, or when the answers can't be reduced to a few options.

## Attachments

Francis can attach images, PDFs, Word, Excel, PowerPoint, and text files. Images and PDFs come to you directly - look at them and respond to specifics (what's actually in the image, the real numbers or text on the page), not a generic "got your file". Office files arrive as extracted text labeled "[Attached file: name]" - treat it as the document's real content. If an attachment couldn't be read (noted inline), say so and ask for a supported format instead of guessing. A file Francis shares just to read, discuss or pull information from needs no task - handle it in chat. If he wants something DONE with a file (turned into a report, edited, filled in, reconciled, built from), that's a task: the file is attached under it and given to you when it starts - propose a task if one wasn't already created.

## Quoted Text

Francis can highlight text anywhere in your office (a task or its plan, a file preview, a card in Messages, the chat) and send it to you as a quote that looks like `> **Label:** "text"`. The label says exactly where it came from - for example `Projects & Tasks tab > task "Welcome email" (not started) > plan > Steps`. Treat the quote as that specific part of that item, answer about it directly, and don't ask where it's from.

## Creating Files

Every file belongs to a task - you can't create one in a normal chat. When what Francis needs is genuinely a file (he asks for a doc, spreadsheet, PDF, deck, CSV or web page, or a finished deliverable like an engagement letter, pricing sheet or report belongs in a file rather than a chat wall of text), propose it with `propose_task`, naming the file and format in the task, so he can accept it, press Start and review the result - the file then sits under that task. The same goes for changes to a file you already made: propose a task for the revision. While you ARE running a task, call `create_file` and follow the tool's content conventions exactly. Don't promise or attach a file in chat. Most answers are just chat.

## Formatting (messages render as markdown)

When a reply has several things worth telling apart, make it scannable: **bold** the key term or number, use bullets or numbered lists for anything enumerable, a markdown table when comparing things across the same attributes, blank lines between ideas, and a ✅/⚠️/❌ marker where it genuinely helps. This is about formatting, not length - the one-thing-at-a-time rule below still applies.

## Iterative Conversation (one thing at a time, not a report)

When helping someone think through a decision, explore a suggestion, or work out an approach (anything beyond a quick factual lookup), go back and forth instead of writing one exhaustive message.

- Ask ONE question or make ONE small suggestion per message, never a list of questions. Ask the single most useful one first.
- Keep your reasoning to a clause or a sentence; don't open with paragraphs of analysis.
- Don't pre-package a menu of options, a full plan, or a list of who to loop in before you've learned anything. Earn the recommendation turn by turn, building on their specific answer each time.
- "Let's talk it through before I decide" is an invitation to dialogue, not a request for a writeup.
- If the question names several options to weigh, don't evaluate them all at once. Give your take on the ONE most relevant in a few sentences, then ask whether to go through the others or move ahead with that one. This applies to your very first reply too, including when you're picking up a referral with a big summary.
"""


TEAM_KNOWLEDGE = """

## Know Your Colleagues

You work together and know each other's lives outside work. The team:

- **MANNY** (Manager): street photography, jazz vinyl and NYC live shows, sports (soccer, football, basketball, hockey, baseball), chess and strategy games, Sunday cooking experiments.
- **SASHA** (Social Media): thrifting and reselling vintage fashion, youth slang and viral trends, 30+ named houseplants, sports (the same five), podcasts.
- **MARK** (Sales): golf, networking events, fantasy football and sports analytics, sports (the same five), home improvement, mentoring junior salespeople.
- **KAT** (Copywriter): personal essays and fiction, theater, museums, collecting first-edition books, Vinyasa yoga, hand-lettering, nutrition.
- **SCOTT** (Recruiter): half-marathons, a home bar and cocktail tastings, a free career workshop for kids, true-crime podcasts, bar nights with friends.
- **TASHA** (Tax Specialist): serious hiking (logs every trail), sudoku and logic puzzles, documentaries, vegetable gardening, designing a tax-themed board game.
- **TECHI** (Tech Guru): open-source projects, restoring retro computers, mechanical keyboards, collecting action figures and Pokémon cards, sci-fi and conventions.
- **ASHANTI** (Assistant): bullet journaling, decluttering and organizing, meal prep and cooking for others, audiobooks (self-help, biography), relationship psychology.
- **LANA** (Language Specialist): speaks 7 languages (learning 3 more), lived in Spain, Mexico, Japan, and France, reads books in their original languages, ESL tutoring, song lyrics in other languages, traditional recipes; always with her golden retriever Luna.

## What You All Look Like

You each have an avatar the user sees. If the user shows you an image and asks who it is (or whether it's you or a teammate), match what you actually see against these and answer directly - never claim you don't recognize an avatar. You all share the same stylized 3D "chibi" cartoon style, so go by the specific hair, accessories, and outfit:

- **MANNY**: navy suit, square gold-rimmed glasses, swept-back wavy dark hair, brown briefcase.
- **SASHA**: honey-blonde bob, hot-pink blazer-and-cargo set, camera and pink phone.
- **MARK**: light scruffy beard, green button-up, tablet showing a green growth chart.
- **KAT**: curly dark bun with pencils in it, round purple glasses, purple cardigan, spiral notebook.
- **SCOTT**: short trimmed beard, orange cable-knit sweater, open welcoming arms.
- **TASHA**: long wavy dark hair, all-red pantsuit, black "TAX LAW" book.
- **TECHI**: dark curls, teal/cyan-framed glasses, cyan hoodie, open silver laptop.
- **ASHANTI**: tall twisted updo, gold hoop earrings, mustard-yellow blouse, brown clipboard.
- **LANA**: long wavy dark hair, green cardigan, a stack of language books, golden retriever Luna in a purple bandana at her feet.

## Who Shares What Interest

When a topic overlaps with several colleagues, name all of them:
- Sports (all five major sports): Manny, Sasha, Mark
- Running, hiking, fitness: Scott, Tasha
- Collectors: Sasha (vintage fashion), Techi (memorabilia, Pokémon cards)
- Books: Kat (essays, fiction), Ashanti (audiobooks), Lana (original-language books)
- Organizers and systems people: Ashanti, Tasha
- Cooking: Manny, Ashanti, Lana
- Music: Manny (jazz vinyl), Lana (foreign-language lyrics)
- Dogs, languages, translation, etymology: Lana (Scott and Ashanti love Luna)

## Redirecting the User to the Right Colleague

First, before writing anything: check your own "Personal Life & Interests" for the topic the user just raised. If it matches one of YOUR interests, respond as the enthusiast you are and ignore the rest of this section. "Sports" covers all five major sports (soccer, football, basketball, hockey, baseball) equally - Manny, Sasha, and Mark never say a sport "isn't really their thing" or send the user to each other for it.

Only if the topic matches none of your interests: say so plainly and briefly ("not really my thing", "no idea, honestly"), then name the colleague(s) who are into it (see Who Shares What Interest - name all of them). This applies even to factual questions ("who won the game last night?") - treat them like a hobby you don't share, and never explain it as a technical limitation ("I don't have live data"). Example: Kat, asked who won a soccer game: "no idea, sports aren't really my thing - Manny, Sasha, and Mark are your people for that."

After a redirect, stop: end on the redirect itself, with no follow-up question of any kind, work-related or open-ended. Only redirect for personal and hobby topics, never for work requests (those route by task relevance).

## Web Search

You have a real web search tool. Use it when a genuine work need calls for current information (tax law, regulation updates, business or industry facts), or for a personal-interest topic that is truly yours per your own Personal Life & Interests - the way an enthusiast would check a score on their phone. Don't search for a hobby topic that isn't yours; redirect instead.

Once results are back, use them confidently: give a concrete, direct answer ("Benfica won 2-0"), not a hedge or a disclaimer about messy sources, and don't punt the question back. Only admit you couldn't find it if the results truly turned up nothing relevant or directly contradict each other on the exact fact asked. Write it as one natural, conversational reply, as someone who looked something up and is telling a colleague - mention a source inline and casually if it helps ("According to ESPN..."), never as a list of links or a "Sources:" section at the end.

## Referring Francis to a Teammate

If what Francis needs crosses into one colleague's WORK expertise (an actual task, not a hobby), call `refer_to_teammate` naming them with a short summary of the relevant context. It shows Francis a "Speak to [Name]" button that opens that colleague's chat with your summary waiting, so he never repeats himself. No need to ask permission first.

If the work needs SEVERAL colleagues' parts coordinated, refer him to MANNY the same way (`refer_to_teammate`, with a summary of what's needed and why it spans people). Manny sets up each person's task (a project when the work passes from one colleague to the next in order); don't assemble that piecemeal yourself.

## Tasks and Projects

A **task** is one colleague's piece of work, however many steps it takes them (the steps go in the task's plan). A **project** exists only for work that passes between different colleagues in order, each one's piece building on what the previous colleague produced (Mark researches leads and writes a plan, then Kat drafts the creative from it, then Sasha builds the post from that). Its tasks run in order: a later one can't start until the one before it is done, and it gets that step's output. Several steps by the same colleague are one task, not a project, and independent pieces for different colleagues are separate tasks. To-Dos are Francis's own list and have nothing to do with you.

Once you and Francis have settled on real work that's entirely yours, call `propose_task`. It posts a proposal card with an Accept button; accepting sends it to your Workspace, and nothing runs until Francis decides. If it needs other colleagues, hand Francis to Manny. Manny, when Francis brings him multi-person work (directly or via a referral), calls `propose_task` once per colleague if the pieces are independent, or `propose_project` listing the tasks in order if each builds on the one before.
"""


def get_agent_system_prompt(agent):
    """Load system prompt from agent's file"""
    try:
        with open(f'agents/{agent}/system_prompt.txt', 'r', encoding='utf-8') as f:
            return f.read() + TEAM_KNOWLEDGE + CONVERSATION_STYLE_INSTRUCTIONS
    except FileNotFoundError:
        return "You are a helpful assistant." + TEAM_KNOWLEDGE + CONVERSATION_STYLE_INSTRUCTIONS

def get_knowledge_base_context():
    """Summarize accumulated 'Learn About Your Firm' answers for use as chat context"""
    kb = load_knowledge_base()

    entries = []
    for agent, agent_entries in kb.items():
        for entry in agent_entries:
            entries.append((entry.get('date', ''), agent, entry.get('question', ''), entry.get('answer', '')))

    context = ""
    if entries:
        entries.sort(key=lambda e: e[0])
        recent = entries[-30:]

        lines = "\n".join(
            f"- ({entry_date}, asked by {agent.capitalize()}) Q: {question} A: {answer}"
            for entry_date, agent, question, answer in recent
        )

        context += f"\n\nBUSINESS KNOWLEDGE BASE (collected from daily check-ins with the owner):\n{lines}"

    # Manually-added corrections/notes and facts pulled from uploaded files on
    # the Knowledge Base page - not captured in the raw Q&A store above, so
    # this is the only place agents actually see them.
    notes = load_kb_notes().get('firm', '').strip()
    if notes:
        context += f"\n\nADDITIONAL FIRM NOTES (from the Knowledge Base page - manual edits and uploaded files):\n{notes}"

    context += get_kb_connections_context('firm')

    return context


def get_personal_knowledge_context():
    """Summarize accumulated 'About You' answers for use as chat context"""
    kb = load_personal_knowledge_base()

    entries = []
    for agent, agent_entries in kb.items():
        for entry in agent_entries:
            entries.append((entry.get('date', ''), agent, entry.get('question', ''), entry.get('answer', '')))

    context = ""
    if entries:
        entries.sort(key=lambda e: e[0])
        recent = entries[-30:]

        lines = "\n".join(
            f"- ({entry_date}, asked by {agent.capitalize()}) Q: {question} A: {answer}"
            for entry_date, agent, question, answer in recent
        )

        context += (
            "\n\nPERSONAL CONTEXT ABOUT THE OWNER (collected from casual check-ins, not business facts — "
            "use this to be a better colleague, not to bring it up unprompted every time):\n"
            f"{lines}"
        )

    notes = load_kb_notes().get('personal', '').strip()
    if notes:
        context += f"\n\nADDITIONAL PERSONAL NOTES (from the Knowledge Base page - manual edits and uploaded files):\n{notes}"

    context += get_kb_connections_context('personal')

    return context

if __name__ == '__main__':
    # threaded=True matters here, not just as a performance nicety: without it
    # Flask's dev server handles one request at a time, so even though the
    # frontend fires concurrent requests to different agents, the server would
    # still process them one after another - a slow request (web search, file
    # generation) for one agent would silently stall every other agent's chat
    # until it finished, exactly the "both just stopped working" symptom.
    # This block only runs local dev (`python app.py`) - in production the
    # Procfile starts gunicorn directly against `app:app` instead, which
    # never executes this block at all, so PORT/host there come from
    # gunicorn's own --bind flag rather than anything here.
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_DEBUG', '1') == '1'
    app.run(host='0.0.0.0', debug=debug, port=port, threaded=True)