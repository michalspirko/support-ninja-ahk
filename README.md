# Support Ninja AHK

A dual-purpose AutoHotkey productivity powerhouse combining lightning-fast text expansion with comprehensive keyboard shortcut reference overlays.

## Key Features

### 🚀 Smart Text Expansion
- Type `+keyword` to instantly insert pre-configured responses
- Dynamic placeholders with custom input dialogs (`{{username:User Name}}`)
- Context-aware formatting (HTML for Teams/Outlook, plain text elsewhere)
- Works seamlessly in Microsoft Teams, web forms, and any text input area
- Perfect for help desk, IT support, and repetitive communications

### ⌨️ Keyboard Shortcuts Overlay
- Visual reference for 200+ shortcuts across Windows 11, Office, browsers
- Organized by application (Teams, Outlook, Chrome, PowerToys, etc.)
- Searchable categories with clean, readable interface
- Always accessible with `Ctrl+Shift+S`

### 🔧 Easy Configuration
- Simple text files for responses (`responses.conf`) and shortcuts (`shortcuts.conf`)
- Markdown link support with automatic format conversion
- Hot-reload configuration without restarting

## Quick Start

1. Configure templates in `responses.conf` and shortcuts in `shortcuts.conf`
2. Run the script
3. **Text Expansion**: Type `+test` → "Hello World"
4. **Shortcuts Reference**: `Ctrl+Shift+S` → Browse shortcut categories
5. **Response List**: `Ctrl+Shift+L` → View all available templates

## Creating Dynamic Canned Responses

### Adding Variables (Placeholders)

Variables are prompted for input when you insert the canned response. Use double curly braces:

```ini
# Simple variable (uses variable name as prompt)
greeting=Hello {{name}}, how can I help you today?

# Custom label (shows "Full Name" in prompt instead of "username")
passwordreset=Hello {{username:Full Name}}, your password reset link is ready.

# Multiple variables
ticket=Hello {{client:Client Name}}, 
Your ticket {{ticketid:Ticket ID}} for {{issue:Issue Description}} has been resolved.
```

When you type `+passwordreset`, a dialog appears asking for "Full Name", then inserts the response with your input.

### Adding Links

Use markdown syntax `[text](url)` - the script automatically formats links appropriately:

- **In Teams/Outlook**: Converts to clickable HTML links
- **In web forms/ServiceNow**: Converts to "text (url)" format

```ini
support=Please visit our [knowledge base](https://kb.company.com) or contact [IT support](mailto:it@company.com).
```

### Multi-line Responses

Use triple quotes for longer responses:

```ini
vpnsetup="""
Hello {{name:User Name}},

To set up VPN on {{device:Device Type}}:

1. Download from [VPN portal](https://vpn.company.com)
2. Use server: {{server:Server Address}}
3. Contact [help desk](mailto:help@company.com) if issues persist

Ticket: {{ticket:Ticket Number}}
"""
```

## Hotkeys

- `Ctrl+Shift+L` - Show response list
- `Ctrl+Shift+S` - Show shortcut categories list
- `Ctrl+Shift+F12` - Reload configuration

## Installation

1. Install [AutoHotkey v2](https://www.autohotkey.com/)
2. Download the script files
3. Configure your responses and shortcuts
4. Run `support-ninja.ahk`

## Credits

Vibe coded by Claude Sonnet 4 ✨

Transform your workflow whether you're answering support tickets or just want quick access to keyboard shortcuts!
