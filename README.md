# Microsoft 365 MCP Server

ABBI's Microsoft 365 integration via Microsoft Graph API. Provides email, calendar, and user management capabilities.

## Features

### Email Tools
- `read_emails` - Read emails from inbox with filtering
- `get_email` - Get full email content by ID  
- `send_email` - Send emails from any authorized mailbox
- `reply_email` - Reply or reply-all to messages
- `search_emails` - Search with KQL syntax

### Calendar Tools
- `list_calendar_events` - List events for date range
- `create_event` - Create events with optional Teams meeting
- `get_availability` - Check free/busy for users

### User Tools
- `list_users` - List organization users
- `get_user` - Get user profile details

## Authentication

Uses Azure AD Application permissions (client credentials flow):
- `Mail.ReadWrite` - Read/write mail in all mailboxes
- `Mail.Send` - Send mail as any user
- `Calendars.ReadWrite` - Read/write calendars
- `User.Read.All` - Read user profiles

## Configuration

Environment variables:
- `TENANT_ID` - Azure AD tenant ID
- `CLIENT_ID` - Azure AD app client ID
- `CLIENT_SECRET` - Azure AD app client secret
- `DEFAULT_USER` - Default mailbox (John.Claude@middleground.com)

## Deployment

Deployed as Azure Container App in SovereignMind-RG resource group.

Container: `m365-mcp.lemoncoast-87756bcf.eastus.azurecontainerapps.io`

## Part of Sovereign Mind

This MCP server integrates with the unified gateway at `sm-mcp-gateway` for seamless tool access across the ABBI ecosystem.
