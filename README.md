# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()

**Turn your running Outlook Desktop into an MCP server with 29 tools.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, manage your calendar, create tasks, handle attachments, and more — all through your existing Outlook session.

## Quick Start

**1. Configure MCP Server (using uvx)**

To use `outlook-desktop-mcp`, add the following snippet to your MCP client's config file (e.g., `claude_desktop_config.json`). This uses `uvx` to automatically install and run the server (requires Python 3.12+ on Windows):

```json
{
  "mcpServers": {
    "outlook-desktop": {
      "command": "uvx",
      "args": [
        "outlook-desktop-mcp"
      ]
    }
  }
}
```


## Requirements

- **Windows** — COM automation is Windows-only
- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

## Available Tools (29)

All tool descriptions are optimized for LLM tool discovery — Claude understands exactly how to use each one, what arguments to pass, and what to expect back.

### 💡 Important Configurations

Many tools support the following dynamic properties for advanced control:

- **`encoding`** *(default: "utf-8")*: Used in `read_email`, `save_attachment`, etc. Pass an encoding string (e.g., `"euc-kr"`, `"cp949"`) alongside your tool call to correctly fetch non-English characters in email subjects, bodies, or file names.
- **`folder_depth` / `max_depth`** *(default: 4)*: Outlook MCP searches for your target folders recursively up to this limit to resolve names into objects. If you hide folders deep within your hierarchy, you can bump this parameter (e.g. `10`) when calling tools like `list_folders`, `list_emails`, or `move_email`.

### Email (9 tools)

| Tool | Description |
|------|-------------|
| `send_email` | Send an email with To/CC/BCC, plain text or HTML body |
| `list_emails` | List recent emails from any folder, with optional unread filter |
| `read_email` | Read full email content by entry ID or subject search |
| `search_emails` | Full-text search across email subjects and bodies |
| `reply_email` | Reply or reply-all, preserving the conversation thread |
| `mark_as_read` | Mark a specific email as read |
| `mark_as_unread` | Mark a specific email as unread |
| `move_email` | Move an email to Archive, Trash, or any folder |
| `list_folders` | Browse the complete folder hierarchy with item counts |

### Calendar (8 tools)

| Tool | Description |
|------|-------------|
| `list_events` | List upcoming events with recurring occurrence support |
| `get_event` | Read full event details by entry ID |
| `create_event` | Create a personal calendar appointment |
| `create_meeting` | Create a meeting and send invitations to attendees |
| `update_event` | Modify an existing event's subject, time, location, etc. |
| `delete_event` | Delete an appointment or cancel a meeting (sends notices) |
| `respond_to_meeting` | Accept, decline, or tentatively accept a meeting invite |
| `search_events` | Search calendar events by keyword within a date range |

### Tasks (5 tools)

| Tool | Description |
|------|-------------|
| `list_tasks` | List pending or completed tasks, sorted by due date |
| `get_task` | Read full task details including body and completion status |
| `create_task` | Create a new task with subject, due date, importance |
| `complete_task` | Mark a task as complete (100%) |
| `delete_task` | Remove a task |

### Attachments (2 tools)

| Tool | Description |
|------|-------------|
| `list_attachments` | List all attachments on an email or calendar event |
| `save_attachment` | Download an attachment to a local directory |

### Categories (2 tools)

| Tool | Description |
|------|-------------|
| `list_categories` | List all available color categories in Outlook |
| `set_category` | Set or clear categories on any email, event, or task |

### Rules (2 tools)

| Tool | Description |
|------|-------------|
| `list_rules` | List all mail rules with enabled/disabled status |
| `toggle_rule` | Enable or disable a mail rule by name |

### Out of Office (1 tool)

| Tool | Description |
|------|-------------|
| `get_out_of_office` | Check whether Out of Office auto-reply is on or off |

## License

See [LICENSE](LICENSE) file.
