# outlook-mail-mcp

Read-only queries over one person's Outlook mailbox through Microsoft Graph. The same queries are available from a CLI and from an MCP server.

## Language

**Operation**:
A named mailbox action such as `search_emails` or `get_email`. It owns its input shape, defaults, limits and description, so the CLI and the MCP server expose it with no rules of their own.
_Avoid_: tool, command, endpoint

**Mailbox**:
The whole store of one person's Outlook mail, with its folders. The project reads it and never changes it.
_Avoid_: inbox, account

**Email**:
One message in the **Mailbox**. Operation names use this word, for example `search_emails`.
_Avoid_: mail, item
