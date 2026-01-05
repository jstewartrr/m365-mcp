"""
Microsoft 365 MCP Server v2 - Full Email Management
Provides comprehensive email, calendar, and user management via Microsoft Graph API
"""

from flask import Flask, request, jsonify
import httpx
import os
import json
from datetime import datetime, timedelta
from functools import lru_cache
import time

app = Flask(__name__)

# Configuration
TENANT_ID = os.environ.get("TENANT_ID", "5e558f98-613b-4c55-80e7-4fd3273d8df3")
CLIENT_ID = os.environ.get("CLIENT_ID", "9969e8ef-8c3b-4ea1-bf85-9a88e1371ab4")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
DEFAULT_USER = os.environ.get("DEFAULT_USER", "John.Claude@middleground.com")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

# Token cache
_token_cache = {"token": None, "expires_at": 0}


def get_access_token():
    """Get or refresh access token"""
    if _token_cache["token"] and time.time() < _token_cache["expires_at"] - 60:
        return _token_cache["token"]
    
    response = httpx.post(TOKEN_URL, data={
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    })
    
    data = response.json()
    if "access_token" in data:
        _token_cache["token"] = data["access_token"]
        _token_cache["expires_at"] = time.time() + data.get("expires_in", 3600)
        return data["access_token"]
    else:
        raise Exception(f"Token error: {data}")


def graph_request(method, endpoint, user=None, **kwargs):
    """Make authenticated Graph API request"""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # Prepend user path if needed
    if user and not endpoint.startswith("/users/"):
        endpoint = f"/users/{user}{endpoint}"
    elif not user and not endpoint.startswith("/users/"):
        endpoint = f"/users/{DEFAULT_USER}{endpoint}"
    
    url = f"{GRAPH_BASE}{endpoint}"
    
    with httpx.Client(timeout=30) as client:
        response = client.request(method, url, headers=headers, **kwargs)
        if response.status_code == 204:
            return {"success": True}
        return response.json()


# ============ MCP PROTOCOL ============

TOOLS = [
    # === EMAIL READING ===
    {
        "name": "read_emails",
        "description": "Read emails from inbox. Returns subject, from, date, and preview.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox (default: John.Claude@middleground.com)"},
                "folder": {"type": "string", "description": "Folder name or ID (default: inbox)"},
                "top": {"type": "integer", "description": "Number of emails to return (default: 10)"},
                "unread_only": {"type": "boolean", "description": "Only return unread emails"},
                "flagged_only": {"type": "boolean", "description": "Only return flagged emails"},
                "search": {"type": "string", "description": "Search query"}
            }
        }
    },
    {
        "name": "get_email",
        "description": "Get full email content by ID",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_id": {"type": "string", "description": "Email message ID"}
            },
            "required": ["message_id"]
        }
    },
    {
        "name": "search_emails",
        "description": "Search emails across mailbox",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "query": {"type": "string", "description": "Search query (KQL syntax)"},
                "top": {"type": "integer", "description": "Max results (default: 25)"}
            },
            "required": ["query"]
        }
    },
    
    # === EMAIL SENDING ===
    {
        "name": "send_email",
        "description": "Send an email from the specified user",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "Send as this user (default: John.Claude@middleground.com)"},
                "to": {"type": "array", "items": {"type": "string"}, "description": "Recipient email addresses"},
                "cc": {"type": "array", "items": {"type": "string"}, "description": "CC recipients"},
                "bcc": {"type": "array", "items": {"type": "string"}, "description": "BCC recipients"},
                "subject": {"type": "string", "description": "Email subject"},
                "body": {"type": "string", "description": "Email body (HTML or plain text)"},
                "is_html": {"type": "boolean", "description": "Body is HTML (default: false)"},
                "importance": {"type": "string", "enum": ["low", "normal", "high"], "description": "Email importance"}
            },
            "required": ["to", "subject", "body"]
        }
    },
    {
        "name": "reply_email",
        "description": "Reply to an email",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_id": {"type": "string", "description": "Original message ID"},
                "body": {"type": "string", "description": "Reply body"},
                "reply_all": {"type": "boolean", "description": "Reply to all recipients"}
            },
            "required": ["message_id", "body"]
        }
    },
    {
        "name": "forward_email",
        "description": "Forward an email to other recipients",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_id": {"type": "string", "description": "Message ID to forward"},
                "to": {"type": "array", "items": {"type": "string"}, "description": "Recipients to forward to"},
                "comment": {"type": "string", "description": "Optional comment to add"}
            },
            "required": ["message_id", "to"]
        }
    },
    
    # === EMAIL MANAGEMENT ===
    {
        "name": "flag_email",
        "description": "Flag or unflag an email for follow-up",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_id": {"type": "string", "description": "Email message ID"},
                "flag_status": {"type": "string", "enum": ["flagged", "complete", "notFlagged"], "description": "Flag status"}
            },
            "required": ["message_id", "flag_status"]
        }
    },
    {
        "name": "mark_read",
        "description": "Mark email(s) as read or unread",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_ids": {"type": "array", "items": {"type": "string"}, "description": "Email message IDs"},
                "is_read": {"type": "boolean", "description": "True to mark as read, False for unread"}
            },
            "required": ["message_ids", "is_read"]
        }
    },
    {
        "name": "move_email",
        "description": "Move email(s) to a different folder",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_ids": {"type": "array", "items": {"type": "string"}, "description": "Email message IDs to move"},
                "destination_folder": {"type": "string", "description": "Destination folder name or ID (e.g., 'archive', 'deleteditems', or folder ID)"}
            },
            "required": ["message_ids", "destination_folder"]
        }
    },
    {
        "name": "delete_email",
        "description": "Delete email(s) - moves to Deleted Items or permanently deletes",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_ids": {"type": "array", "items": {"type": "string"}, "description": "Email message IDs to delete"},
                "permanent": {"type": "boolean", "description": "Permanently delete (skip Deleted Items)"}
            },
            "required": ["message_ids"]
        }
    },
    {
        "name": "copy_email",
        "description": "Copy email(s) to another folder",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "message_ids": {"type": "array", "items": {"type": "string"}, "description": "Email message IDs to copy"},
                "destination_folder": {"type": "string", "description": "Destination folder name or ID"}
            },
            "required": ["message_ids", "destination_folder"]
        }
    },
    
    # === FOLDER MANAGEMENT ===
    {
        "name": "list_folders",
        "description": "List all mail folders",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "include_hidden": {"type": "boolean", "description": "Include hidden folders"}
            }
        }
    },
    {
        "name": "create_folder",
        "description": "Create a new mail folder",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "name": {"type": "string", "description": "Folder name"},
                "parent_folder": {"type": "string", "description": "Parent folder ID (optional, default is root)"}
            },
            "required": ["name"]
        }
    },
    {
        "name": "delete_folder",
        "description": "Delete a mail folder",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "folder_id": {"type": "string", "description": "Folder ID to delete"}
            },
            "required": ["folder_id"]
        }
    },
    {
        "name": "rename_folder",
        "description": "Rename a mail folder",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "folder_id": {"type": "string", "description": "Folder ID"},
                "new_name": {"type": "string", "description": "New folder name"}
            },
            "required": ["folder_id", "new_name"]
        }
    },
    
    # === MAIL RULES ===
    {
        "name": "list_rules",
        "description": "List all inbox rules",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"}
            }
        }
    },
    {
        "name": "create_rule",
        "description": "Create an inbox rule",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "name": {"type": "string", "description": "Rule name"},
                "conditions": {"type": "object", "description": "Rule conditions (e.g., from_addresses, subject_contains)"},
                "actions": {"type": "object", "description": "Rule actions (e.g., move_to_folder, mark_as_read, flag)"},
                "enabled": {"type": "boolean", "description": "Enable rule (default: true)"}
            },
            "required": ["name", "conditions", "actions"]
        }
    },
    {
        "name": "update_rule",
        "description": "Update an existing inbox rule",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "rule_id": {"type": "string", "description": "Rule ID to update"},
                "name": {"type": "string", "description": "New rule name"},
                "conditions": {"type": "object", "description": "New conditions"},
                "actions": {"type": "object", "description": "New actions"},
                "enabled": {"type": "boolean", "description": "Enable/disable rule"}
            },
            "required": ["rule_id"]
        }
    },
    {
        "name": "delete_rule",
        "description": "Delete an inbox rule",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User mailbox"},
                "rule_id": {"type": "string", "description": "Rule ID to delete"}
            },
            "required": ["rule_id"]
        }
    },
    
    # === CALENDAR ===
    {
        "name": "list_calendar_events",
        "description": "List calendar events for a date range",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User calendar"},
                "start_date": {"type": "string", "description": "Start date (ISO format, default: today)"},
                "end_date": {"type": "string", "description": "End date (ISO format, default: 7 days from start)"},
                "top": {"type": "integer", "description": "Max events (default: 50)"}
            }
        }
    },
    {
        "name": "create_event",
        "description": "Create a calendar event",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User calendar"},
                "subject": {"type": "string", "description": "Event title"},
                "start": {"type": "string", "description": "Start datetime (ISO format)"},
                "end": {"type": "string", "description": "End datetime (ISO format)"},
                "attendees": {"type": "array", "items": {"type": "string"}, "description": "Attendee emails"},
                "location": {"type": "string", "description": "Location"},
                "body": {"type": "string", "description": "Event description"},
                "is_online": {"type": "boolean", "description": "Create Teams meeting"}
            },
            "required": ["subject", "start", "end"]
        }
    },
    {
        "name": "update_event",
        "description": "Update a calendar event",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User calendar"},
                "event_id": {"type": "string", "description": "Event ID"},
                "subject": {"type": "string", "description": "New title"},
                "start": {"type": "string", "description": "New start datetime"},
                "end": {"type": "string", "description": "New end datetime"},
                "location": {"type": "string", "description": "New location"},
                "body": {"type": "string", "description": "New description"}
            },
            "required": ["event_id"]
        }
    },
    {
        "name": "delete_event",
        "description": "Delete a calendar event",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User calendar"},
                "event_id": {"type": "string", "description": "Event ID to delete"}
            },
            "required": ["event_id"]
        }
    },
    {
        "name": "get_availability",
        "description": "Check free/busy availability for users",
        "inputSchema": {
            "type": "object",
            "properties": {
                "emails": {"type": "array", "items": {"type": "string"}, "description": "User emails to check"},
                "start": {"type": "string", "description": "Start datetime (ISO)"},
                "end": {"type": "string", "description": "End datetime (ISO)"}
            },
            "required": ["emails", "start", "end"]
        }
    },
    
    # === USERS ===
    {
        "name": "list_users",
        "description": "List users in the organization",
        "inputSchema": {
            "type": "object",
            "properties": {
                "search": {"type": "string", "description": "Search by name or email"},
                "top": {"type": "integer", "description": "Max results (default: 50)"}
            }
        }
    },
    {
        "name": "get_user",
        "description": "Get user profile details",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user": {"type": "string", "description": "User email or ID"}
            },
            "required": ["user"]
        }
    }
]


# ============ TOOL IMPLEMENTATIONS ============

def read_emails(user=None, folder="inbox", top=10, unread_only=False, flagged_only=False, search=None):
    """Read emails from inbox"""
    user = user or DEFAULT_USER
    params = [f"$top={top}", "$select=id,subject,from,receivedDateTime,bodyPreview,isRead,flag,importance"]
    params.append("$orderby=receivedDateTime desc")
    
    filters = []
    if unread_only:
        filters.append("isRead eq false")
    if flagged_only:
        filters.append("flag/flagStatus eq 'flagged'")
    if search:
        params.append(f"$search=\"{search}\"")
    if filters:
        params.append(f"$filter={' and '.join(filters)}")
    
    endpoint = f"/mailFolders/{folder}/messages?{'&'.join(params)}"
    result = graph_request("GET", endpoint, user=user)
    
    if "value" in result:
        return {
            "count": len(result["value"]),
            "emails": [{
                "id": m["id"],
                "subject": m.get("subject", "(no subject)"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address", "unknown"),
                "from_name": m.get("from", {}).get("emailAddress", {}).get("name", ""),
                "date": m.get("receivedDateTime"),
                "preview": m.get("bodyPreview", "")[:200],
                "is_read": m.get("isRead", False),
                "flag_status": m.get("flag", {}).get("flagStatus", "notFlagged"),
                "importance": m.get("importance", "normal")
            } for m in result["value"]]
        }
    return result


def get_email(message_id, user=None):
    """Get full email by ID"""
    user = user or DEFAULT_USER
    endpoint = f"/messages/{message_id}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,flag,importance,isRead"
    result = graph_request("GET", endpoint, user=user)
    
    if "id" in result:
        return {
            "id": result["id"],
            "subject": result.get("subject"),
            "from": result.get("from", {}).get("emailAddress", {}),
            "to": [r.get("emailAddress", {}) for r in result.get("toRecipients", [])],
            "cc": [r.get("emailAddress", {}) for r in result.get("ccRecipients", [])],
            "date": result.get("receivedDateTime"),
            "body": result.get("body", {}).get("content", ""),
            "has_attachments": result.get("hasAttachments", False),
            "is_read": result.get("isRead", False),
            "flag_status": result.get("flag", {}).get("flagStatus", "notFlagged"),
            "importance": result.get("importance", "normal")
        }
    return result


def send_email(to, subject, body, user=None, cc=None, bcc=None, is_html=False, importance="normal"):
    """Send an email"""
    user = user or DEFAULT_USER
    
    message = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body
            },
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
            "importance": importance
        }
    }
    
    if cc:
        message["message"]["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc]
    if bcc:
        message["message"]["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc]
    
    result = graph_request("POST", "/sendMail", user=user, json=message)
    return {"success": True, "message": f"Email sent to {', '.join(to)}"}


def reply_email(message_id, body, user=None, reply_all=False):
    """Reply to an email"""
    user = user or DEFAULT_USER
    action = "replyAll" if reply_all else "reply"
    
    result = graph_request("POST", f"/messages/{message_id}/{action}", user=user, json={
        "comment": body
    })
    return {"success": True, "action": action}


def forward_email(message_id, to, user=None, comment=None):
    """Forward an email"""
    user = user or DEFAULT_USER
    
    body = {
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to]
    }
    if comment:
        body["comment"] = comment
    
    result = graph_request("POST", f"/messages/{message_id}/forward", user=user, json=body)
    return {"success": True, "message": f"Email forwarded to {', '.join(to)}"}


def search_emails(query, user=None, top=25):
    """Search emails"""
    user = user or DEFAULT_USER
    endpoint = f"/messages?$search=\"{query}\"&$top={top}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,flag"
    result = graph_request("GET", endpoint, user=user)
    
    if "value" in result:
        return {
            "count": len(result["value"]),
            "emails": [{
                "id": m["id"],
                "subject": m.get("subject", "(no subject)"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address", "unknown"),
                "date": m.get("receivedDateTime"),
                "preview": m.get("bodyPreview", "")[:200],
                "is_read": m.get("isRead", False),
                "flag_status": m.get("flag", {}).get("flagStatus", "notFlagged")
            } for m in result["value"]]
        }
    return result


def flag_email(message_id, flag_status, user=None):
    """Flag or unflag an email"""
    user = user or DEFAULT_USER
    
    body = {
        "flag": {
            "flagStatus": flag_status
        }
    }
    
    result = graph_request("PATCH", f"/messages/{message_id}", user=user, json=body)
    return {"success": True, "message_id": message_id, "flag_status": flag_status}


def mark_read(message_ids, is_read, user=None):
    """Mark emails as read or unread"""
    user = user or DEFAULT_USER
    results = []
    
    for msg_id in message_ids:
        result = graph_request("PATCH", f"/messages/{msg_id}", user=user, json={"isRead": is_read})
        results.append({"message_id": msg_id, "success": "error" not in result})
    
    return {
        "success": True,
        "marked_as": "read" if is_read else "unread",
        "count": len(message_ids),
        "results": results
    }


def move_email(message_ids, destination_folder, user=None):
    """Move emails to a folder"""
    user = user or DEFAULT_USER
    
    # Map common folder names to well-known folder names
    folder_map = {
        "inbox": "inbox",
        "archive": "archive", 
        "deleted": "deleteditems",
        "deleteditems": "deleteditems",
        "trash": "deleteditems",
        "drafts": "drafts",
        "sent": "sentitems",
        "sentitems": "sentitems",
        "junk": "junkemail",
        "junkemail": "junkemail",
        "spam": "junkemail"
    }
    
    folder_id = folder_map.get(destination_folder.lower(), destination_folder)
    results = []
    
    for msg_id in message_ids:
        result = graph_request("POST", f"/messages/{msg_id}/move", user=user, json={"destinationId": folder_id})
        results.append({"message_id": msg_id, "success": "id" in result or "success" in result})
    
    return {
        "success": True,
        "moved_to": destination_folder,
        "count": len(message_ids),
        "results": results
    }


def delete_email(message_ids, user=None, permanent=False):
    """Delete emails"""
    user = user or DEFAULT_USER
    results = []
    
    for msg_id in message_ids:
        if permanent:
            result = graph_request("DELETE", f"/messages/{msg_id}", user=user)
        else:
            result = graph_request("POST", f"/messages/{msg_id}/move", user=user, json={"destinationId": "deleteditems"})
        results.append({"message_id": msg_id, "success": True})
    
    return {
        "success": True,
        "deleted": "permanently" if permanent else "moved to trash",
        "count": len(message_ids),
        "results": results
    }


def copy_email(message_ids, destination_folder, user=None):
    """Copy emails to a folder"""
    user = user or DEFAULT_USER
    results = []
    
    for msg_id in message_ids:
        result = graph_request("POST", f"/messages/{msg_id}/copy", user=user, json={"destinationId": destination_folder})
        results.append({"message_id": msg_id, "new_id": result.get("id"), "success": "id" in result})
    
    return {
        "success": True,
        "copied_to": destination_folder,
        "count": len(message_ids),
        "results": results
    }


def list_folders(user=None, include_hidden=False):
    """List mail folders"""
    user = user or DEFAULT_USER
    params = "$select=id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount"
    if not include_hidden:
        params += "&$filter=isHidden eq false"
    
    endpoint = f"/mailFolders?{params}"
    result = graph_request("GET", endpoint, user=user)
    
    if "value" in result:
        return {
            "count": len(result["value"]),
            "folders": [{
                "id": f["id"],
                "name": f.get("displayName"),
                "parent_id": f.get("parentFolderId"),
                "child_count": f.get("childFolderCount", 0),
                "unread_count": f.get("unreadItemCount", 0),
                "total_count": f.get("totalItemCount", 0)
            } for f in result["value"]]
        }
    return result


def create_folder(name, user=None, parent_folder=None):
    """Create a mail folder"""
    user = user or DEFAULT_USER
    
    if parent_folder:
        endpoint = f"/mailFolders/{parent_folder}/childFolders"
    else:
        endpoint = "/mailFolders"
    
    result = graph_request("POST", endpoint, user=user, json={"displayName": name})
    
    if "id" in result:
        return {"success": True, "folder_id": result["id"], "name": result.get("displayName")}
    return result


def delete_folder(folder_id, user=None):
    """Delete a mail folder"""
    user = user or DEFAULT_USER
    result = graph_request("DELETE", f"/mailFolders/{folder_id}", user=user)
    return {"success": True, "folder_id": folder_id}


def rename_folder(folder_id, new_name, user=None):
    """Rename a mail folder"""
    user = user or DEFAULT_USER
    result = graph_request("PATCH", f"/mailFolders/{folder_id}", user=user, json={"displayName": new_name})
    return {"success": True, "folder_id": folder_id, "new_name": new_name}


def list_rules(user=None):
    """List inbox rules"""
    user = user or DEFAULT_USER
    endpoint = "/mailFolders/inbox/messageRules"
    result = graph_request("GET", endpoint, user=user)
    
    if "value" in result:
        return {
            "count": len(result["value"]),
            "rules": [{
                "id": r["id"],
                "name": r.get("displayName"),
                "sequence": r.get("sequence"),
                "enabled": r.get("isEnabled", True),
                "conditions": r.get("conditions", {}),
                "actions": r.get("actions", {})
            } for r in result["value"]]
        }
    return result


def create_rule(name, conditions, actions, user=None, enabled=True):
    """Create an inbox rule"""
    user = user or DEFAULT_USER
    
    # Build conditions
    rule_conditions = {}
    if "from_addresses" in conditions:
        rule_conditions["fromAddresses"] = [{"emailAddress": {"address": addr}} for addr in conditions["from_addresses"]]
    if "subject_contains" in conditions:
        rule_conditions["subjectContains"] = conditions["subject_contains"] if isinstance(conditions["subject_contains"], list) else [conditions["subject_contains"]]
    if "body_contains" in conditions:
        rule_conditions["bodyContains"] = conditions["body_contains"] if isinstance(conditions["body_contains"], list) else [conditions["body_contains"]]
    if "sender_contains" in conditions:
        rule_conditions["senderContains"] = conditions["sender_contains"] if isinstance(conditions["sender_contains"], list) else [conditions["sender_contains"]]
    if "has_attachments" in conditions:
        rule_conditions["hasAttachments"] = conditions["has_attachments"]
    if "importance" in conditions:
        rule_conditions["importance"] = conditions["importance"]
    
    # Build actions
    rule_actions = {}
    if "move_to_folder" in actions:
        rule_actions["moveToFolder"] = actions["move_to_folder"]
    if "copy_to_folder" in actions:
        rule_actions["copyToFolder"] = actions["copy_to_folder"]
    if "mark_as_read" in actions:
        rule_actions["markAsRead"] = actions["mark_as_read"]
    if "flag" in actions:
        rule_actions["flagStatus"] = actions["flag"]
    if "delete" in actions:
        rule_actions["delete"] = actions["delete"]
    if "forward_to" in actions:
        rule_actions["forwardTo"] = [{"emailAddress": {"address": addr}} for addr in actions["forward_to"]]
    if "stop_processing" in actions:
        rule_actions["stopProcessingRules"] = actions["stop_processing"]
    if "mark_importance" in actions:
        rule_actions["markImportance"] = actions["mark_importance"]
    
    body = {
        "displayName": name,
        "isEnabled": enabled,
        "conditions": rule_conditions,
        "actions": rule_actions
    }
    
    endpoint = "/mailFolders/inbox/messageRules"
    result = graph_request("POST", endpoint, user=user, json=body)
    
    if "id" in result:
        return {"success": True, "rule_id": result["id"], "name": result.get("displayName")}
    return result


def update_rule(rule_id, user=None, name=None, conditions=None, actions=None, enabled=None):
    """Update an inbox rule"""
    user = user or DEFAULT_USER
    
    body = {}
    if name:
        body["displayName"] = name
    if enabled is not None:
        body["isEnabled"] = enabled
    if conditions:
        # Same condition mapping as create_rule
        rule_conditions = {}
        if "from_addresses" in conditions:
            rule_conditions["fromAddresses"] = [{"emailAddress": {"address": addr}} for addr in conditions["from_addresses"]]
        if "subject_contains" in conditions:
            rule_conditions["subjectContains"] = conditions["subject_contains"] if isinstance(conditions["subject_contains"], list) else [conditions["subject_contains"]]
        body["conditions"] = rule_conditions
    if actions:
        # Same action mapping as create_rule
        rule_actions = {}
        if "move_to_folder" in actions:
            rule_actions["moveToFolder"] = actions["move_to_folder"]
        if "mark_as_read" in actions:
            rule_actions["markAsRead"] = actions["mark_as_read"]
        body["actions"] = rule_actions
    
    endpoint = f"/mailFolders/inbox/messageRules/{rule_id}"
    result = graph_request("PATCH", endpoint, user=user, json=body)
    return {"success": True, "rule_id": rule_id}


def delete_rule(rule_id, user=None):
    """Delete an inbox rule"""
    user = user or DEFAULT_USER
    endpoint = f"/mailFolders/inbox/messageRules/{rule_id}"
    result = graph_request("DELETE", endpoint, user=user)
    return {"success": True, "rule_id": rule_id}


def list_calendar_events(user=None, start_date=None, end_date=None, top=50):
    """List calendar events"""
    user = user or DEFAULT_USER
    
    if not start_date:
        start_date = datetime.utcnow().strftime("%Y-%m-%dT00:00:00Z")
    if not end_date:
        end_dt = datetime.utcnow() + timedelta(days=7)
        end_date = end_dt.strftime("%Y-%m-%dT23:59:59Z")
    
    endpoint = f"/calendarView?startDateTime={start_date}&endDateTime={end_date}&$top={top}&$select=id,subject,start,end,location,attendees,isOnlineMeeting,onlineMeetingUrl"
    result = graph_request("GET", endpoint, user=user)
    
    if "value" in result:
        return {
            "count": len(result["value"]),
            "events": [{
                "id": e["id"],
                "subject": e.get("subject"),
                "start": e.get("start", {}).get("dateTime"),
                "end": e.get("end", {}).get("dateTime"),
                "location": e.get("location", {}).get("displayName"),
                "attendees": [a.get("emailAddress", {}).get("address") for a in e.get("attendees", [])],
                "is_online": e.get("isOnlineMeeting", False),
                "teams_url": e.get("onlineMeetingUrl")
            } for e in result["value"]]
        }
    return result


def create_event(subject, start, end, user=None, attendees=None, location=None, body=None, is_online=False):
    """Create calendar event"""
    user = user or DEFAULT_USER
    
    event = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": "UTC"},
        "end": {"dateTime": end, "timeZone": "UTC"}
    }
    
    if attendees:
        event["attendees"] = [{"emailAddress": {"address": addr}, "type": "required"} for addr in attendees]
    if location:
        event["location"] = {"displayName": location}
    if body:
        event["body"] = {"contentType": "Text", "content": body}
    if is_online:
        event["isOnlineMeeting"] = True
        event["onlineMeetingProvider"] = "teamsForBusiness"
    
    result = graph_request("POST", "/events", user=user, json=event)
    
    if "id" in result:
        return {
            "success": True,
            "event_id": result["id"],
            "subject": result.get("subject"),
            "teams_url": result.get("onlineMeeting", {}).get("joinUrl") if is_online else None
        }
    return result


def update_event(event_id, user=None, subject=None, start=None, end=None, location=None, body=None):
    """Update a calendar event"""
    user = user or DEFAULT_USER
    
    update = {}
    if subject:
        update["subject"] = subject
    if start:
        update["start"] = {"dateTime": start, "timeZone": "UTC"}
    if end:
        update["end"] = {"dateTime": end, "timeZone": "UTC"}
    if location:
        update["location"] = {"displayName": location}
    if body:
        update["body"] = {"contentType": "Text", "content": body}
    
    result = graph_request("PATCH", f"/events/{event_id}", user=user, json=update)
    return {"success": True, "event_id": event_id}


def delete_event(event_id, user=None):
    """Delete a calendar event"""
    user = user or DEFAULT_USER
    result = graph_request("DELETE", f"/events/{event_id}", user=user)
    return {"success": True, "event_id": event_id}


def get_availability(emails, start, end):
    """Check free/busy availability"""
    token = get_access_token()
    
    body = {
        "schedules": emails,
        "startTime": {"dateTime": start, "timeZone": "UTC"},
        "endTime": {"dateTime": end, "timeZone": "UTC"},
        "availabilityViewInterval": 30
    }
    
    with httpx.Client(timeout=30) as client:
        response = client.post(
            f"{GRAPH_BASE}/users/{DEFAULT_USER}/calendar/getSchedule",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=body
        )
        result = response.json()
    
    if "value" in result:
        return {
            "schedules": [{
                "email": s.get("scheduleId"),
                "availability": s.get("availabilityView"),
                "busy_slots": [{
                    "start": slot.get("start", {}).get("dateTime"),
                    "end": slot.get("end", {}).get("dateTime"),
                    "status": slot.get("status")
                } for slot in s.get("scheduleItems", [])]
            } for s in result["value"]]
        }
    return result


def list_users(search=None, top=50):
    """List organization users"""
    token = get_access_token()
    
    params = [f"$top={top}", "$select=id,displayName,mail,jobTitle,department"]
    if search:
        params.append(f"$search=\"displayName:{search}\" OR \"mail:{search}\"")
    
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    if search:
        headers["ConsistencyLevel"] = "eventual"
    
    with httpx.Client(timeout=30) as client:
        response = client.get(f"{GRAPH_BASE}/users?{'&'.join(params)}", headers=headers)
        result = response.json()
    
    if "value" in result:
        return {
            "count": len(result["value"]),
            "users": [{
                "id": u["id"],
                "name": u.get("displayName"),
                "email": u.get("mail"),
                "title": u.get("jobTitle"),
                "department": u.get("department")
            } for u in result["value"] if u.get("mail")]
        }
    return result


def get_user(user):
    """Get user profile"""
    token = get_access_token()
    
    with httpx.Client(timeout=30) as client:
        response = client.get(
            f"{GRAPH_BASE}/users/{user}?$select=id,displayName,mail,jobTitle,department,officeLocation,mobilePhone,businessPhones",
            headers={"Authorization": f"Bearer {token}"}
        )
        result = response.json()
    
    if "id" in result:
        return {
            "id": result["id"],
            "name": result.get("displayName"),
            "email": result.get("mail"),
            "title": result.get("jobTitle"),
            "department": result.get("department"),
            "office": result.get("officeLocation"),
            "mobile": result.get("mobilePhone"),
            "phone": result.get("businessPhones", [None])[0]
        }
    return result


# Tool dispatcher
TOOL_MAP = {
    "read_emails": read_emails,
    "get_email": get_email,
    "send_email": send_email,
    "reply_email": reply_email,
    "forward_email": forward_email,
    "search_emails": search_emails,
    "flag_email": flag_email,
    "mark_read": mark_read,
    "move_email": move_email,
    "delete_email": delete_email,
    "copy_email": copy_email,
    "list_folders": list_folders,
    "create_folder": create_folder,
    "delete_folder": delete_folder,
    "rename_folder": rename_folder,
    "list_rules": list_rules,
    "create_rule": create_rule,
    "update_rule": update_rule,
    "delete_rule": delete_rule,
    "list_calendar_events": list_calendar_events,
    "create_event": create_event,
    "update_event": update_event,
    "delete_event": delete_event,
    "get_availability": get_availability,
    "list_users": list_users,
    "get_user": get_user
}


# ============ ROUTES ============

@app.route("/health", methods=["GET"])
def health():
    """Health check"""
    try:
        token = get_access_token()
        return jsonify({
            "status": "healthy",
            "service": "m365-mcp",
            "version": "2.0.0",
            "authenticated": True,
            "default_user": DEFAULT_USER,
            "tools": len(TOOLS)
        })
    except Exception as e:
        return jsonify({"status": "unhealthy", "error": str(e)}), 500


@app.route("/mcp", methods=["POST"])
def mcp_handler():
    """MCP protocol handler"""
    data = request.json or {}
    method = data.get("method")
    
    if method == "initialize":
        return jsonify({
            "jsonrpc": "2.0",
            "id": data.get("id"),
            "result": {
                "protocolVersion": "2024-11-05",
                "serverInfo": {"name": "m365-mcp", "version": "2.0.0"},
                "capabilities": {"tools": {"listChanged": False}}
            }
        })
    
    elif method == "tools/list":
        return jsonify({
            "jsonrpc": "2.0",
            "id": data.get("id"),
            "result": {"tools": TOOLS}
        })
    
    elif method == "tools/call":
        tool_name = data.get("params", {}).get("name")
        arguments = data.get("params", {}).get("arguments", {})
        
        if tool_name not in TOOL_MAP:
            return jsonify({
                "jsonrpc": "2.0",
                "id": data.get("id"),
                "error": {"code": -32601, "message": f"Unknown tool: {tool_name}"}
            })
        
        try:
            result = TOOL_MAP[tool_name](**arguments)
            return jsonify({
                "jsonrpc": "2.0",
                "id": data.get("id"),
                "result": {"content": [{"type": "text", "text": json.dumps(result, indent=2, default=str)}]}
            })
        except Exception as e:
            return jsonify({
                "jsonrpc": "2.0",
                "id": data.get("id"),
                "error": {"code": -32000, "message": str(e)}
            })
    
    return jsonify({
        "jsonrpc": "2.0",
        "id": data.get("id"),
        "error": {"code": -32601, "message": f"Method not found: {method}"}
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
