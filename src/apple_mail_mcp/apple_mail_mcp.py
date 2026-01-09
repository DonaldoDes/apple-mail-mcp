#!/usr/bin/env python3
"""
Apple Mail MCP Server - FastMCP implementation
Provides tools to query and interact with Apple Mail inboxes
"""

import asyncio
import subprocess
import json
import os
import logging
from datetime import datetime
from typing import Optional, List, Dict, Any
from mcp.server.fastmcp import FastMCP

# Configure logging
logger = logging.getLogger(__name__)

# Load user preferences from environment
USER_PREFERENCES = os.environ.get("USER_EMAIL_PREFERENCES", "")

# Initialize FastMCP server
mcp = FastMCP("Apple Mail MCP")

# Global lock to serialize AppleScript executions
# Apple Mail can only handle one operation at a time
_applescript_lock = asyncio.Lock()

# Retry configuration
MAX_RETRIES = 3
INITIAL_BACKOFF_SECONDS = 2


def get_inbox_applescript_snippet(account_var: str = "anAccount") -> str:
    """
    Returns AppleScript code snippet for dynamic inbox discovery.
    
    Supports multiple inbox name variations across different mail providers:
    - IMAP standard: "INBOX", "Inbox"
    - Exchange/Outlook (localized): "Boîte de réception" (FR), "Posteingang" (DE), etc.
    
    Args:
        account_var: The AppleScript variable name for the account (default: "anAccount")
    
    Returns:
        AppleScript code that sets 'inboxMailbox' variable to the discovered inbox
    """
    return f'''
                -- Dynamic inbox discovery with fallback chain
                set inboxMailbox to missing value
                set possibleInboxNames to {{"INBOX", "Inbox", "Boîte de réception", "Posteingang", "Bandeja de entrada", "Posta in arrivo", "Caixa de entrada", "Входящие", "受信トレイ", "收件箱"}}
                repeat with inboxName in possibleInboxNames
                    try
                        set inboxMailbox to mailbox inboxName of {account_var}
                        exit repeat
                    end try
                end repeat
                if inboxMailbox is missing value then
                    error "Could not find inbox for account " & (name of {account_var})
                end if
'''


def get_inbox_applescript_snippet_for_target(account_var: str = "targetAccount") -> str:
    """
    Returns AppleScript code snippet for dynamic inbox discovery for a target account.
    
    This is a convenience wrapper that uses 'targetAccount' as the default variable name,
    which is commonly used when accessing a specific account by name.
    
    Args:
        account_var: The AppleScript variable name for the account (default: "targetAccount")
    
    Returns:
        AppleScript code that sets 'inboxMailbox' variable to the discovered inbox
    """
    return get_inbox_applescript_snippet(account_var)


def get_mailbox_applescript_snippet(mailbox: str, account_var: str = "targetAccount", result_var: str = "searchMailbox") -> str:
    """
    Returns AppleScript code snippet for mailbox discovery with INBOX fallback.
    
    When the mailbox is "INBOX", uses dynamic discovery to find the inbox
    across different mail providers (Exchange, IMAP, etc.).
    
    Args:
        mailbox: The mailbox name to find
        account_var: The AppleScript variable name for the account (default: "targetAccount")
        result_var: The AppleScript variable name for the result (default: "searchMailbox")
    
    Returns:
        AppleScript code that sets the result_var to the discovered mailbox
    """
    if mailbox.upper() == "INBOX":
        # Use dynamic inbox discovery for INBOX
        return f'''
            -- Dynamic inbox discovery with fallback chain
            set {result_var} to missing value
            set possibleInboxNames to {{"INBOX", "Inbox", "Boîte de réception", "Posteingang", "Bandeja de entrada", "Posta in arrivo", "Caixa de entrada", "Входящие", "受信トレイ", "收件箱"}}
            repeat with inboxName in possibleInboxNames
                try
                    set {result_var} to mailbox inboxName of {account_var}
                    exit repeat
                end try
            end repeat
            if {result_var} is missing value then
                error "Could not find inbox for account"
            end if
'''
    else:
        # For other mailboxes, try direct access
        return f'''
            set {result_var} to mailbox "{mailbox}" of {account_var}
'''


# Decorator to inject user preferences into tool docstrings
def inject_preferences(func):
    """Decorator that appends user preferences to tool docstrings"""
    if USER_PREFERENCES:
        if func.__doc__:
            func.__doc__ = func.__doc__.rstrip() + f"\n\nUser Preferences: {USER_PREFERENCES}"
        else:
            func.__doc__ = f"User Preferences: {USER_PREFERENCES}"
    return func


async def run_applescript(script: str) -> str:
    """Execute AppleScript and return output with retry logic.
    
    Uses asyncio.to_thread() to run the blocking subprocess call
    in a thread pool, making it non-blocking for the async event loop.
    
    A global lock ensures only one AppleScript runs at a time,
    as Apple Mail cannot handle concurrent operations.
    
    Retry behavior:
    - Max 3 attempts with exponential backoff (2s, 4s, 8s)
    - Only retries on timeout errors
    - Script errors (syntax, runtime) fail immediately
    """
    async with _applescript_lock:
        last_exception = None
        
        for attempt in range(MAX_RETRIES):
            try:
                logger.debug(f"AppleScript execution attempt {attempt + 1}/{MAX_RETRIES}")
                
                result = await asyncio.to_thread(
                    subprocess.run,
                    ['osascript', '-e', script],
                    capture_output=True,
                    text=True,
                    timeout=120
                )
                
                # Check for AppleScript errors
                if result.returncode != 0:
                    error_msg = result.stderr.strip() if result.stderr else "Unknown AppleScript error"
                    # Script errors should not be retried - they will fail every time
                    raise Exception(f"AppleScript error (code {result.returncode}): {error_msg}")
                
                if attempt > 0:
                    logger.info(f"AppleScript succeeded on attempt {attempt + 1}")
                
                return result.stdout.strip()
                
            except subprocess.TimeoutExpired as e:
                last_exception = e
                backoff_time = INITIAL_BACKOFF_SECONDS * (2 ** attempt)
                logger.warning(
                    f"AppleScript timeout on attempt {attempt + 1}/{MAX_RETRIES}. "
                    f"Retrying in {backoff_time}s..."
                )
                
                if attempt < MAX_RETRIES - 1:
                    await asyncio.sleep(backoff_time)
                else:
                    raise Exception(
                        f"AppleScript execution timed out after {MAX_RETRIES} attempts. "
                        "Apple Mail may be unresponsive."
                    )
                    
            except FileNotFoundError:
                # This error won't be fixed by retrying
                raise Exception("osascript not found. This tool requires macOS with AppleScript support.")
                
            except Exception as e:
                if "AppleScript error" in str(e):
                    # Script errors should not be retried
                    raise
                # For other unexpected errors, don't retry
                raise Exception(f"AppleScript execution failed: {str(e)}")
        
        # Should not reach here, but just in case
        raise Exception(f"AppleScript execution failed after {MAX_RETRIES} attempts: {str(last_exception)}")


def parse_email_list(output: str) -> List[Dict[str, Any]]:
    """Parse the structured email output from AppleScript"""
    emails = []
    lines = output.split('\n')

    current_email = {}
    for line in lines:
        line = line.strip()
        if not line or line.startswith('=') or line.startswith('━') or line.startswith('📧') or line.startswith('⚠'):
            continue

        if line.startswith('✉') or line.startswith('✓'):
            # New email entry
            if current_email:
                emails.append(current_email)

            is_read = line.startswith('✓')
            subject = line[2:].strip()  # Remove indicator
            current_email = {
                'subject': subject,
                'is_read': is_read
            }
        elif line.startswith('From:'):
            current_email['sender'] = line[5:].strip()
        elif line.startswith('Date:'):
            current_email['date'] = line[5:].strip()
        elif line.startswith('Preview:'):
            current_email['preview'] = line[8:].strip()
        elif line.startswith('TOTAL EMAILS'):
            # End of email list
            if current_email:
                emails.append(current_email)
            break

    if current_email and current_email not in emails:
        emails.append(current_email)

    return emails


@mcp.tool()
@inject_preferences
async def list_inbox_emails(
    account: Optional[str] = None,
    max_emails: int = 0,
    include_read: bool = True
) -> str:
    """
    List all emails from inbox across all accounts or a specific account.

    Args:
        account: Optional account name to filter (e.g., "Gmail", "Work"). If None, shows all accounts.
        max_emails: Maximum number of emails to return per account (0 = all)
        include_read: Whether to include read emails (default: True)

    Returns:
        Formatted list of emails with subject, sender, date, and read status
    """

    inbox_discovery = get_inbox_applescript_snippet("anAccount")
    
    script = f'''
    tell application "Mail"
        set outputText to "INBOX EMAILS - ALL ACCOUNTS" & return & return
        set totalCount to 0
        set allAccounts to every account

        repeat with anAccount in allAccounts
            set accountName to name of anAccount

            try
                {inbox_discovery}
                set inboxMessages to every message of inboxMailbox
                set messageCount to count of inboxMessages

                if messageCount > 0 then
                    set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
                    set outputText to outputText & "📧 ACCOUNT: " & accountName & " (" & messageCount & " messages)" & return
                    set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

                    set currentIndex to 0
                    repeat with aMessage in inboxMessages
                        set currentIndex to currentIndex + 1
                        if {max_emails} > 0 and currentIndex > {max_emails} then exit repeat

                        try
                            set messageSubject to subject of aMessage
                            set messageSender to sender of aMessage
                            set messageDate to date received of aMessage
                            set messageRead to read status of aMessage

                            set shouldInclude to true
                            if not {str(include_read).lower()} and messageRead then
                                set shouldInclude to false
                            end if

                            if shouldInclude then
                                if messageRead then
                                    set readIndicator to "✓"
                                else
                                    set readIndicator to "✉"
                                end if

                                set outputText to outputText & readIndicator & " " & messageSubject & return
                                set outputText to outputText & "   From: " & messageSender & return
                                set outputText to outputText & "   Date: " & (messageDate as string) & return
                                set outputText to outputText & return

                                set totalCount to totalCount + 1
                            end if
                        end try
                    end repeat
                end if
            on error errMsg
                set outputText to outputText & "⚠ Error accessing inbox for account " & accountName & return
                set outputText to outputText & "   " & errMsg & return & return
            end try
        end repeat

        set outputText to outputText & "========================================" & return
        set outputText to outputText & "TOTAL EMAILS: " & totalCount & return
        set outputText to outputText & "========================================" & return

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def get_email_with_content(
    account: str,
    subject_keyword: str,
    max_results: int = 5,
    max_content_length: int = 300,
    mailbox: str = "INBOX"
) -> str:
    """
    Search for emails by subject keyword and return with full content preview.

    Args:
        account: Account name to search in (e.g., "Gmail", "Work")
        subject_keyword: Keyword to search for in email subjects
        max_results: Maximum number of matching emails to return (default: 5)
        max_content_length: Maximum content length in characters (default: 300, 0 = unlimited)
        mailbox: Mailbox to search (default: "INBOX", use "All" for all mailboxes)

    Returns:
        Detailed email information including content preview
    """

    # Build mailbox selection logic
    if mailbox == "All":
        mailbox_script = '''
            set allMailboxes to every mailbox of targetAccount
            set searchMailboxes to allMailboxes
        '''
        search_location = "all mailboxes"
    else:
        mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "searchMailbox")
        mailbox_script = f'''
            try
                {mailbox_discovery}
            on error errMsg
                error "Mailbox not found: {mailbox}. " & errMsg
            end try
            set searchMailboxes to {{searchMailbox}}
        '''
        search_location = mailbox

    script = f'''
    on lowercase(str)
        set lowerStr to do shell script "echo " & quoted form of str & " | tr '[:upper:]' '[:lower:]'"
        return lowerStr
    end lowercase

    tell application "Mail"
        set outputText to "SEARCH RESULTS FOR: {subject_keyword}" & return
        set outputText to outputText & "Searching in: {search_location}" & return & return
        set resultCount to 0

        try
            set targetAccount to account "{account}"
            {mailbox_script}

            repeat with currentMailbox in searchMailboxes
                set mailboxMessages to every message of currentMailbox
                set mailboxName to name of currentMailbox

                repeat with aMessage in mailboxMessages
                    if resultCount >= {max_results} then exit repeat

                    try
                        set messageSubject to subject of aMessage

                        -- Convert to lowercase for case-insensitive matching
                        set lowerSubject to my lowercase(messageSubject)
                        set lowerKeyword to my lowercase("{subject_keyword}")

                        -- Check if subject contains keyword (case insensitive)
                        if lowerSubject contains lowerKeyword then
                            set messageSender to sender of aMessage
                            set messageDate to date received of aMessage
                            set messageRead to read status of aMessage

                            if messageRead then
                                set readIndicator to "✓"
                            else
                                set readIndicator to "✉"
                            end if

                            set outputText to outputText & readIndicator & " " & messageSubject & return
                            set outputText to outputText & "   From: " & messageSender & return
                            set outputText to outputText & "   Date: " & (messageDate as string) & return
                            set outputText to outputText & "   Mailbox: " & mailboxName & return

                            -- Get content preview
                            try
                                set msgContent to content of aMessage
                                set AppleScript's text item delimiters to {{return, linefeed}}
                                set contentParts to text items of msgContent
                                set AppleScript's text item delimiters to " "
                                set cleanText to contentParts as string
                                set AppleScript's text item delimiters to ""

                                -- Handle content length limit (0 = unlimited)
                                if {max_content_length} > 0 and length of cleanText > {max_content_length} then
                                    set contentPreview to text 1 thru {max_content_length} of cleanText & "..."
                                else
                                    set contentPreview to cleanText
                                end if

                                set outputText to outputText & "   Content: " & contentPreview & return
                            on error
                                set outputText to outputText & "   Content: [Not available]" & return
                            end try

                            set outputText to outputText & return
                            set resultCount to resultCount + 1
                        end if
                    end try
                end repeat
            end repeat

            set outputText to outputText & "========================================" & return
            set outputText to outputText & "FOUND: " & resultCount & " matching email(s)" & return
            set outputText to outputText & "========================================" & return

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def get_unread_count() -> Dict[str, int]:
    """
    Get the count of unread emails for each account.

    Returns:
        Dictionary mapping account names to unread email counts
    """

    inbox_discovery = get_inbox_applescript_snippet("anAccount")
    
    script = f'''
    tell application "Mail"
        set resultList to {{}}
        set allAccounts to every account

        repeat with anAccount in allAccounts
            set accountName to name of anAccount

            try
                {inbox_discovery}
                set unreadCount to unread count of inboxMailbox
                set end of resultList to accountName & ":" & unreadCount
            on error
                set end of resultList to accountName & ":ERROR"
            end try
        end repeat

        set AppleScript's text item delimiters to "|"
        return resultList as string
    end tell
    '''

    result = await run_applescript(script)

    # Parse the result
    counts = {}
    for item in result.split('|'):
        if ':' in item:
            account, count = item.split(':', 1)
            if count != "ERROR":
                counts[account] = int(count)
            else:
                counts[account] = -1  # Error indicator

    return counts


@mcp.tool()
@inject_preferences
async def list_accounts() -> List[str]:
    """
    List all available Mail accounts.

    Returns:
        List of account names
    """

    script = '''
    tell application "Mail"
        set accountNames to {}
        set allAccounts to every account

        repeat with anAccount in allAccounts
            set accountName to name of anAccount
            set end of accountNames to accountName
        end repeat

        set AppleScript's text item delimiters to "|"
        return accountNames as string
    end tell
    '''

    result = await run_applescript(script)
    return result.split('|') if result else []


@mcp.tool()
@inject_preferences
async def get_recent_emails(
    account: str,
    count: int = 10,
    include_content: bool = False
) -> str:
    """
    Get the most recent emails from a specific account.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        count: Number of recent emails to retrieve (default: 10)
        include_content: Whether to include content preview (slower, default: False)

    Returns:
        Formatted list of recent emails
    """

    content_script = '''
        try
            set msgContent to content of aMessage
            set AppleScript's text item delimiters to {{return, linefeed}}
            set contentParts to text items of msgContent
            set AppleScript's text item delimiters to " "
            set cleanText to contentParts as string
            set AppleScript's text item delimiters to ""

            if length of cleanText > 200 then
                set contentPreview to text 1 thru 200 of cleanText & "..."
            else
                set contentPreview to cleanText
            end if

            set outputText to outputText & "   Preview: " & contentPreview & return
        on error
            set outputText to outputText & "   Preview: [Not available]" & return
        end try
    ''' if include_content else ''

    inbox_discovery = get_inbox_applescript_snippet_for_target("targetAccount")
    
    script = f'''
    tell application "Mail"
        set outputText to "RECENT EMAILS - {account}" & return & return

        try
            set targetAccount to account "{account}"
            {inbox_discovery}
            set inboxMessages to every message of inboxMailbox

            set currentIndex to 0
            repeat with aMessage in inboxMessages
                set currentIndex to currentIndex + 1
                if currentIndex > {count} then exit repeat

                try
                    set messageSubject to subject of aMessage
                    set messageSender to sender of aMessage
                    set messageDate to date received of aMessage
                    set messageRead to read status of aMessage

                    if messageRead then
                        set readIndicator to "✓"
                    else
                        set readIndicator to "✉"
                    end if

                    set outputText to outputText & readIndicator & " " & messageSubject & return
                    set outputText to outputText & "   From: " & messageSender & return
                    set outputText to outputText & "   Date: " & (messageDate as string) & return

                    {content_script}

                    set outputText to outputText & return
                end try
            end repeat

            set outputText to outputText & "========================================" & return
            set outputText to outputText & "Showing " & (currentIndex - 1) & " email(s)" & return
            set outputText to outputText & "========================================" & return

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def list_mailboxes(
    account: Optional[str] = None,
    include_counts: bool = True
) -> str:
    """
    List all mailboxes (folders) for a specific account or all accounts.

    Args:
        account: Optional account name to filter (e.g., "Gmail", "Work"). If None, shows all accounts.
        include_counts: Whether to include message counts for each mailbox (default: True)

    Returns:
        Formatted list of mailboxes with optional message counts.
        For nested mailboxes, shows both indented format and path format (e.g., "Projects/Amplify Impact")
    """

    count_script = '''
        try
            set msgCount to count of messages of aMailbox
            set unreadCount to unread count of aMailbox
            set outputText to outputText & " (" & msgCount & " total, " & unreadCount & " unread)"
        on error
            set outputText to outputText & " (count unavailable)"
        end try
    ''' if include_counts else ''

    account_filter = f'''
        if accountName is "{account}" then
    ''' if account else ''

    account_filter_end = 'end if' if account else ''

    script = f'''
    tell application "Mail"
        set outputText to "MAILBOXES" & return & return
        set allAccounts to every account

        repeat with anAccount in allAccounts
            set accountName to name of anAccount

            {account_filter}
                set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
                set outputText to outputText & "📁 ACCOUNT: " & accountName & return
                set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

                try
                    set accountMailboxes to every mailbox of anAccount

                    repeat with aMailbox in accountMailboxes
                        set mailboxName to name of aMailbox
                        set outputText to outputText & "  📂 " & mailboxName

                        {count_script}

                        set outputText to outputText & return

                        -- List sub-mailboxes with path notation
                        try
                            set subMailboxes to every mailbox of aMailbox
                            repeat with subBox in subMailboxes
                                set subName to name of subBox
                                set outputText to outputText & "    └─ " & subName & " [Path: " & mailboxName & "/" & subName & "]"

                                {count_script.replace('aMailbox', 'subBox') if include_counts else ''}

                                set outputText to outputText & return
                            end repeat
                        end try
                    end repeat

                    set outputText to outputText & return
                on error errMsg
                    set outputText to outputText & "  ⚠ Error accessing mailboxes: " & errMsg & return & return
                end try
            {account_filter_end}
        end repeat

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def move_email(
    account: str,
    subject_keyword: str,
    to_mailbox: str,
    from_mailbox: str = "INBOX",
    max_moves: int = 1
) -> str:
    """
    Move email(s) matching a subject keyword from one mailbox to another.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        subject_keyword: Keyword to search for in email subjects
        to_mailbox: Destination mailbox name. For nested mailboxes, use "/" separator (e.g., "Projects/Amplify Impact")
        from_mailbox: Source mailbox name (default: "INBOX")
        max_moves: Maximum number of emails to move (default: 1, safety limit)

    Returns:
        Confirmation message with details of moved emails
    """

    # Parse nested mailbox path
    mailbox_parts = to_mailbox.split('/')

    # Build the nested mailbox reference
    if len(mailbox_parts) > 1:
        # Nested mailbox
        dest_mailbox_script = f'mailbox "{mailbox_parts[-1]}" of '
        for i in range(len(mailbox_parts) - 2, -1, -1):
            dest_mailbox_script += f'mailbox "{mailbox_parts[i]}" of '
        dest_mailbox_script += 'targetAccount'
    else:
        # Top-level mailbox
        dest_mailbox_script = f'mailbox "{to_mailbox}" of targetAccount'

    source_mailbox_discovery = get_mailbox_applescript_snippet(from_mailbox, "targetAccount", "sourceMailbox")
    
    script = f'''
    tell application "Mail"
        set outputText to "MOVING EMAILS" & return & return
        set movedCount to 0

        try
            set targetAccount to account "{account}"
            -- Get source mailbox with dynamic inbox discovery
            try
                {source_mailbox_discovery}
            on error errMsg
                error "Source mailbox not found: {from_mailbox}. " & errMsg
            end try

            -- Get destination mailbox (handles nested mailboxes)
            set destMailbox to {dest_mailbox_script}
            set sourceMessages to every message of sourceMailbox

            repeat with aMessage in sourceMessages
                if movedCount >= {max_moves} then exit repeat

                try
                    set messageSubject to subject of aMessage

                    -- Check if subject contains keyword (case insensitive)
                    if messageSubject contains "{subject_keyword}" then
                        set messageSender to sender of aMessage
                        set messageDate to date received of aMessage

                        -- Move the message
                        move aMessage to destMailbox

                        set outputText to outputText & "✓ Moved: " & messageSubject & return
                        set outputText to outputText & "  From: " & messageSender & return
                        set outputText to outputText & "  Date: " & (messageDate as string) & return
                        set outputText to outputText & "  {from_mailbox} → {to_mailbox}" & return & return

                        set movedCount to movedCount + 1
                    end if
                end try
            end repeat

            set outputText to outputText & "========================================" & return
            set outputText to outputText & "TOTAL MOVED: " & movedCount & " email(s)" & return
            set outputText to outputText & "========================================" & return

        on error errMsg
            return "Error: " & errMsg & return & "Please check that account and mailbox names are correct. For nested mailboxes, use '/' separator (e.g., 'Projects/Amplify Impact')."
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def reply_to_email(
    account: str,
    subject_keyword: str,
    reply_body: str,
    reply_to_all: bool = False,
    confirm: bool = False
) -> str:
    """
    Reply to an email matching a subject keyword.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        subject_keyword: Keyword to search for in email subjects
        reply_body: The body text of the reply
        reply_to_all: If True, reply to all recipients; if False, reply only to sender (default: False)
        confirm: If False (default), shows preview without sending.
                 If True, actually sends the reply.

    Returns:
        Confirmation message with details of the reply (preview or sent)
    """

    # Escape quotes in reply_body for AppleScript
    escaped_body = reply_body.replace('"', '\\"')

    # Build the reply command based on reply_to_all flag
    if reply_to_all:
        reply_command = 'set replyMessage to reply foundMessage with opening window reply to all'
    else:
        reply_command = 'set replyMessage to reply foundMessage with opening window'

    # Send command based on confirm flag
    if confirm:
        send_command = 'send replyMessage'
        status_message = "✓ Reply sent successfully!"
    else:
        send_command = '-- send replyMessage (dry run - set confirm=True to send)'
        status_message = "📋 PREVIEW - Reply prepared but NOT sent (set confirm=True to send)"

    inbox_discovery = get_inbox_applescript_snippet_for_target("targetAccount")
    
    script = f'''
    tell application "Mail"
        set outputText to "SENDING REPLY" & return & return

        try
            set targetAccount to account "{account}"
            {inbox_discovery}
            set inboxMessages to every message of inboxMailbox
            set foundMessage to missing value

            -- Find the first matching message
            repeat with aMessage in inboxMessages
                try
                    set messageSubject to subject of aMessage

                    if messageSubject contains "{subject_keyword}" then
                        set foundMessage to aMessage
                        exit repeat
                    end if
                end try
            end repeat

            if foundMessage is not missing value then
                set messageSubject to subject of foundMessage
                set messageSender to sender of foundMessage
                set messageDate to date received of foundMessage

                -- Create reply
                {reply_command}

                -- Ensure the reply is from the correct account
                set sender of replyMessage to targetAccount

                -- Set reply content
                set content of replyMessage to "{escaped_body}"

                -- Send the reply (or not, based on confirm)
                {send_command}

                set outputText to outputText & "{status_message}" & return & return
                set outputText to outputText & "Original email:" & return
                set outputText to outputText & "  Subject: " & messageSubject & return
                set outputText to outputText & "  From: " & messageSender & return
                set outputText to outputText & "  Date: " & (messageDate as string) & return & return
                set outputText to outputText & "Reply body:" & return
                set outputText to outputText & "  " & "{escaped_body}" & return

            else
                set outputText to outputText & "⚠ No email found matching: {subject_keyword}" & return
            end if

        on error errMsg
            return "Error: " & errMsg & return & "Please check that the account name is correct and the email exists."
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def compose_email(
    account: str,
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    confirm: bool = False
) -> str:
    """
    Compose and send a new email from a specific account.

    Args:
        account: Account name to send from (e.g., "Gmail", "Work", "Personal")
        to: Recipient email address(es), comma-separated for multiple
        subject: Email subject line
        body: Email body text
        cc: Optional CC recipients, comma-separated for multiple
        bcc: Optional BCC recipients, comma-separated for multiple
        confirm: If False (default), shows preview without sending.
                 If True, actually sends the email.

    Returns:
        Confirmation message with details of the email (preview or sent)
    """

    # Escape quotes for AppleScript
    escaped_subject = subject.replace('"', '\\"')
    escaped_body = body.replace('"', '\\"')

    # Build CC recipients if provided
    cc_script = ''
    if cc:
        cc_addresses = [addr.strip() for addr in cc.split(',')]
        for addr in cc_addresses:
            cc_script += f'''
            make new cc recipient at end of cc recipients of newMessage with properties {{address:"{addr}"}}
            '''

    # Build BCC recipients if provided
    bcc_script = ''
    if bcc:
        bcc_addresses = [addr.strip() for addr in bcc.split(',')]
        for addr in bcc_addresses:
            bcc_script += f'''
            make new bcc recipient at end of bcc recipients of newMessage with properties {{address:"{addr}"}}
            '''

    # Send command based on confirm flag
    if confirm:
        send_command = 'send newMessage'
        status_message = "✓ Email sent successfully!"
    else:
        send_command = '-- send newMessage (dry run - set confirm=True to send)'
        status_message = "📋 PREVIEW - Email prepared but NOT sent (set confirm=True to send)"

    script = f'''
    tell application "Mail"
        set outputText to "COMPOSING EMAIL" & return & return

        try
            set targetAccount to account "{account}"

            -- Create new outgoing message
            set newMessage to make new outgoing message with properties {{subject:"{escaped_subject}", content:"{escaped_body}", visible:false}}

            -- Set the sender account
            set sender of newMessage to targetAccount

            -- Add TO recipients
            tell newMessage
                make new to recipient at end of to recipients with properties {{address:"{to}"}}
                {cc_script}
                {bcc_script}
            end tell

            -- Send the message (or not, based on confirm)
            {send_command}

            set outputText to outputText & "{status_message}" & return & return
            set outputText to outputText & "From: " & name of targetAccount & return
            set outputText to outputText & "To: {to}" & return
    '''

    if cc:
        script += f'''
            set outputText to outputText & "CC: {cc}" & return
    '''

    if bcc:
        script += f'''
            set outputText to outputText & "BCC: {bcc}" & return
    '''

    script += f'''
            set outputText to outputText & "Subject: {escaped_subject}" & return
            set outputText to outputText & "Body: " & "{escaped_body}" & return

        on error errMsg
            return "Error: " & errMsg & return & "Please check that the account name and email addresses are correct."
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def list_email_attachments(
    account: str,
    subject_keyword: str,
    max_results: int = 1
) -> str:
    """
    List attachments for emails matching a subject keyword.

    Args:
        account: Account name (e.g., "Gmail", "Work", "Personal")
        subject_keyword: Keyword to search for in email subjects
        max_results: Maximum number of matching emails to check (default: 1)

    Returns:
        List of attachments with their names and sizes
    """

    inbox_discovery = get_inbox_applescript_snippet_for_target("targetAccount")
    
    script = f'''
    tell application "Mail"
        set outputText to "ATTACHMENTS FOR: {subject_keyword}" & return & return
        set resultCount to 0

        try
            set targetAccount to account "{account}"
            {inbox_discovery}
            set inboxMessages to every message of inboxMailbox

            repeat with aMessage in inboxMessages
                if resultCount >= {max_results} then exit repeat

                try
                    set messageSubject to subject of aMessage

                    -- Check if subject contains keyword
                    if messageSubject contains "{subject_keyword}" then
                        set messageSender to sender of aMessage
                        set messageDate to date received of aMessage

                        set outputText to outputText & "✉ " & messageSubject & return
                        set outputText to outputText & "   From: " & messageSender & return
                        set outputText to outputText & "   Date: " & (messageDate as string) & return & return

                        -- Get attachments
                        set msgAttachments to mail attachments of aMessage
                        set attachmentCount to count of msgAttachments

                        if attachmentCount > 0 then
                            set outputText to outputText & "   Attachments (" & attachmentCount & "):" & return

                            repeat with anAttachment in msgAttachments
                                set attachmentName to name of anAttachment
                                try
                                    set attachmentSize to size of anAttachment
                                    set sizeInKB to (attachmentSize / 1024) as integer
                                    set outputText to outputText & "   📎 " & attachmentName & " (" & sizeInKB & " KB)" & return
                                on error
                                    set outputText to outputText & "   📎 " & attachmentName & return
                                end try
                            end repeat
                        else
                            set outputText to outputText & "   No attachments" & return
                        end if

                        set outputText to outputText & return
                        set resultCount to resultCount + 1
                    end if
                end try
            end repeat

            set outputText to outputText & "========================================" & return
            set outputText to outputText & "FOUND: " & resultCount & " matching email(s)" & return
            set outputText to outputText & "========================================" & return

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def save_email_attachment(
    account: str,
    subject_keyword: str,
    attachment_name: str,
    save_path: str
) -> str:
    """
    Save a specific attachment from an email to disk.

    Args:
        account: Account name (e.g., "Gmail", "Work", "Personal")
        subject_keyword: Keyword to search for in email subjects
        attachment_name: Name of the attachment to save
        save_path: Full path where to save the attachment

    Returns:
        Confirmation message with save location
    """

    inbox_discovery = get_inbox_applescript_snippet_for_target("targetAccount")
    
    script = f'''
    tell application "Mail"
        set outputText to ""

        try
            set targetAccount to account "{account}"
            {inbox_discovery}
            set inboxMessages to every message of inboxMailbox
            set foundAttachment to false

            repeat with aMessage in inboxMessages
                try
                    set messageSubject to subject of aMessage

                    -- Check if subject contains keyword
                    if messageSubject contains "{subject_keyword}" then
                        set msgAttachments to mail attachments of aMessage

                        repeat with anAttachment in msgAttachments
                            set attachmentFileName to name of anAttachment

                            if attachmentFileName contains "{attachment_name}" then
                                -- Save the attachment
                                save anAttachment in POSIX file "{save_path}"

                                set outputText to "✓ Attachment saved successfully!" & return & return
                                set outputText to outputText & "Email: " & messageSubject & return
                                set outputText to outputText & "Attachment: " & attachmentFileName & return
                                set outputText to outputText & "Saved to: {save_path}" & return

                                set foundAttachment to true
                                exit repeat
                            end if
                        end repeat

                        if foundAttachment then exit repeat
                    end if
                end try
            end repeat

            if not foundAttachment then
                set outputText to "⚠ Attachment not found" & return
                set outputText to outputText & "Email keyword: {subject_keyword}" & return
                set outputText to outputText & "Attachment name: {attachment_name}" & return
            end if

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def get_inbox_overview() -> str:
    """
    Get a comprehensive overview of your email inbox status across all accounts.

    Returns:
        Comprehensive overview including:
        - Unread email counts by account
        - List of available mailboxes/folders
        - AI suggestions for actions (move emails, respond to messages, highlight action items, etc.)

    This tool is designed to give you a complete picture of your inbox and prompt the assistant
    to suggest relevant actions based on the current state.
    """

    inbox_discovery = get_inbox_applescript_snippet("anAccount")
    
    script = f'''
    tell application "Mail"
        set outputText to "╔══════════════════════════════════════════╗" & return
        set outputText to outputText & "║      EMAIL INBOX OVERVIEW                ║" & return
        set outputText to outputText & "╚══════════════════════════════════════════╝" & return & return

        -- Section 1: Unread Counts by Account
        set outputText to outputText & "📊 UNREAD EMAILS BY ACCOUNT" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
        set allAccounts to every account
        set totalUnread to 0

        repeat with anAccount in allAccounts
            set accountName to name of anAccount

            try
                {inbox_discovery}

                set unreadCount to unread count of inboxMailbox
                set totalMessages to count of messages of inboxMailbox
                set totalUnread to totalUnread + unreadCount

                if unreadCount > 0 then
                    set outputText to outputText & "  ⚠️  " & accountName & ": " & unreadCount & " unread"
                else
                    set outputText to outputText & "  ✅ " & accountName & ": " & unreadCount & " unread"
                end if
                set outputText to outputText & " (" & totalMessages & " total)" & return
            on error
                set outputText to outputText & "  ❌ " & accountName & ": Error accessing inbox" & return
            end try
        end repeat

        set outputText to outputText & return
        set outputText to outputText & "📈 TOTAL UNREAD: " & totalUnread & " across all accounts" & return
        set outputText to outputText & return & return

        -- Section 2: Mailboxes/Folders Overview
        set outputText to outputText & "📁 MAILBOX STRUCTURE" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return

        repeat with anAccount in allAccounts
            set accountName to name of anAccount
            set outputText to outputText & return & "Account: " & accountName & return

            try
                set accountMailboxes to every mailbox of anAccount

                repeat with aMailbox in accountMailboxes
                    set mailboxName to name of aMailbox

                    try
                        set unreadCount to unread count of aMailbox
                        if unreadCount > 0 then
                            set outputText to outputText & "  📂 " & mailboxName & " (" & unreadCount & " unread)" & return
                        else
                            set outputText to outputText & "  📂 " & mailboxName & return
                        end if

                        -- Show nested mailboxes if they have unread messages
                        try
                            set subMailboxes to every mailbox of aMailbox
                            repeat with subBox in subMailboxes
                                set subName to name of subBox
                                set subUnread to unread count of subBox

                                if subUnread > 0 then
                                    set outputText to outputText & "     └─ " & subName & " (" & subUnread & " unread)" & return
                                end if
                            end repeat
                        end try
                    on error
                        set outputText to outputText & "  📂 " & mailboxName & return
                    end try
                end repeat
            on error
                set outputText to outputText & "  ⚠️  Error accessing mailboxes" & return
            end try
        end repeat

        set outputText to outputText & return & return

        -- Section 3: Recent Emails Preview (10 most recent across all accounts)
        set outputText to outputText & "📬 RECENT EMAILS PREVIEW (10 Most Recent)" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return

        -- Collect all recent messages from all accounts
        set allRecentMessages to {{}}

        repeat with anAccount in allAccounts
            set accountName to name of anAccount

            try
                {inbox_discovery}

                set inboxMessages to every message of inboxMailbox

                -- Get up to 10 messages from each account
                set messageIndex to 0
                repeat with aMessage in inboxMessages
                    set messageIndex to messageIndex + 1
                    if messageIndex > 10 then exit repeat

                    try
                        set messageSubject to subject of aMessage
                        set messageSender to sender of aMessage
                        set messageDate to date received of aMessage
                        set messageRead to read status of aMessage

                        -- Create message record
                        set messageRecord to {{accountName:accountName, msgSubject:messageSubject, msgSender:messageSender, msgDate:messageDate, msgRead:messageRead}}
                        set end of allRecentMessages to messageRecord
                    end try
                end repeat
            end try
        end repeat

        -- Display up to 10 most recent messages
        set displayCount to 0
        repeat with msgRecord in allRecentMessages
            set displayCount to displayCount + 1
            if displayCount > 10 then exit repeat

            set readIndicator to "✉"
            if msgRead of msgRecord then
                set readIndicator to "✓"
            end if

            set outputText to outputText & return & readIndicator & " " & msgSubject of msgRecord & return
            set outputText to outputText & "   Account: " & accountName of msgRecord & return
            set outputText to outputText & "   From: " & msgSender of msgRecord & return
            set outputText to outputText & "   Date: " & (msgDate of msgRecord as string) & return
        end repeat

        if displayCount = 0 then
            set outputText to outputText & return & "No recent emails found." & return
        end if

        set outputText to outputText & return & return

        -- Section 4: Action Suggestions (for the AI assistant)
        set outputText to outputText & "💡 SUGGESTED ACTIONS FOR ASSISTANT" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
        set outputText to outputText & "Based on this overview, consider suggesting:" & return & return

        if totalUnread > 0 then
            set outputText to outputText & "1. 📧 Review unread emails - Use get_recent_emails() to show recent unread messages" & return
            set outputText to outputText & "2. 🔍 Search for action items - Look for keywords like 'urgent', 'action required', 'deadline'" & return
            set outputText to outputText & "3. 📤 Move processed emails - Suggest moving read emails to appropriate folders" & return
        else
            set outputText to outputText & "1. ✅ Inbox is clear! No unread emails." & return
        end if

        set outputText to outputText & "4. 📋 Organize by topic - Suggest moving emails to project-specific folders" & return
        set outputText to outputText & "5. ✉️  Draft replies - Identify emails that need responses" & return
        set outputText to outputText & "6. 🗂️  Archive old emails - Move older read emails to archive folders" & return
        set outputText to outputText & "7. 🔔 Highlight priority items - Identify emails from important senders or with urgent keywords" & return

        set outputText to outputText & return
        set outputText to outputText & "═══════════════════════════════════════════════════" & return
        set outputText to outputText & "💬 Ask me to drill down into any account or take specific actions!" & return
        set outputText to outputText & "═══════════════════════════════════════════════════" & return

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def search_emails(
    account: str,
    mailbox: str = "INBOX",
    subject_keyword: Optional[str] = None,
    sender: Optional[str] = None,
    has_attachments: Optional[bool] = None,
    read_status: str = "all",
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    include_content: bool = False,
    max_results: int = 20
) -> str:
    """
    Unified search tool - search emails with advanced filtering across any mailbox.

    Args:
        account: Account name to search in (e.g., "Gmail", "Work")
        mailbox: Mailbox to search (default: "INBOX", use "All" for all mailboxes, or specific folder name)
        subject_keyword: Optional keyword to search in subject
        sender: Optional sender email or name to filter by
        has_attachments: Optional filter for emails with attachments (True/False/None)
        read_status: Filter by read status: "all", "read", "unread" (default: "all")
        date_from: Optional start date filter (format: "YYYY-MM-DD")
        date_to: Optional end date filter (format: "YYYY-MM-DD")
        include_content: Whether to include email content preview (slower)
        max_results: Maximum number of results to return (default: 20)

    Returns:
        Formatted list of matching emails with all requested details
    """

    # Build AppleScript search conditions
    conditions = []

    if subject_keyword:
        conditions.append(f'messageSubject contains "{subject_keyword}"')

    if sender:
        conditions.append(f'messageSender contains "{sender}"')

    if has_attachments is not None:
        if has_attachments:
            conditions.append('(count of mail attachments of aMessage) > 0')
        else:
            conditions.append('(count of mail attachments of aMessage) = 0')

    if read_status == "read":
        conditions.append('messageRead is true')
    elif read_status == "unread":
        conditions.append('messageRead is false')

    # Combine conditions with AND logic
    condition_str = ' and '.join(conditions) if conditions else 'true'

    # Handle content preview
    content_script = '''
        try
            set msgContent to content of aMessage
            set AppleScript's text item delimiters to {{return, linefeed}}
            set contentParts to text items of msgContent
            set AppleScript's text item delimiters to " "
            set cleanText to contentParts as string
            set AppleScript's text item delimiters to ""

            if length of cleanText > 300 then
                set contentPreview to text 1 thru 300 of cleanText & "..."
            else
                set contentPreview to cleanText
            end if

            set outputText to outputText & "   Content: " & contentPreview & return
        on error
            set outputText to outputText & "   Content: [Not available]" & return
        end try
    ''' if include_content else ''

    # Build mailbox selection logic
    if mailbox == "All":
        mailbox_script = '''
            set allMailboxes to every mailbox of targetAccount
            set searchMailboxes to allMailboxes
        '''
    else:
        mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "searchMailbox")
        mailbox_script = f'''
            try
                {mailbox_discovery}
            on error errMsg
                error "Mailbox not found: {mailbox}. " & errMsg
            end try
            set searchMailboxes to {{searchMailbox}}
        '''

    script = f'''
    tell application "Mail"
        set outputText to "SEARCH RESULTS" & return & return
        set outputText to outputText & "Searching in: {mailbox}" & return
        set outputText to outputText & "Account: {account}" & return & return
        set resultCount to 0

        try
            set targetAccount to account "{account}"
            {mailbox_script}

            repeat with currentMailbox in searchMailboxes
                set mailboxMessages to every message of currentMailbox
                set mailboxName to name of currentMailbox

                repeat with aMessage in mailboxMessages
                    if resultCount >= {max_results} then exit repeat

                    try
                        set messageSubject to subject of aMessage
                        set messageSender to sender of aMessage
                        set messageDate to date received of aMessage
                        set messageRead to read status of aMessage

                        -- Apply search conditions
                        if {condition_str} then
                            set readIndicator to "✉"
                            if messageRead then
                                set readIndicator to "✓"
                            end if

                            set outputText to outputText & readIndicator & " " & messageSubject & return
                            set outputText to outputText & "   From: " & messageSender & return
                            set outputText to outputText & "   Date: " & (messageDate as string) & return
                            set outputText to outputText & "   Mailbox: " & mailboxName & return

                            {content_script}

                            set outputText to outputText & return
                            set resultCount to resultCount + 1
                        end if
                    end try
                end repeat
            end repeat

            set outputText to outputText & "========================================" & return
            set outputText to outputText & "FOUND: " & resultCount & " matching email(s)" & return
            set outputText to outputText & "========================================" & return

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def update_email_status(
    account: str,
    action: str,
    subject_keyword: Optional[str] = None,
    sender: Optional[str] = None,
    mailbox: str = "INBOX",
    max_updates: int = 10
) -> str:
    """
    Update email status - mark as read/unread or flag/unflag emails.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        action: Action to perform: "mark_read", "mark_unread", "flag", "unflag"
        subject_keyword: Optional keyword to filter emails by subject
        sender: Optional sender to filter emails by
        mailbox: Mailbox to search in (default: "INBOX")
        max_updates: Maximum number of emails to update (safety limit, default: 10)

    Returns:
        Confirmation message with details of updated emails
    """

    # Build search condition
    conditions = []
    if subject_keyword:
        conditions.append(f'messageSubject contains "{subject_keyword}"')
    if sender:
        conditions.append(f'messageSender contains "{sender}"')

    condition_str = ' and '.join(conditions) if conditions else 'true'

    # Build action script
    if action == "mark_read":
        action_script = 'set read status of aMessage to true'
        action_label = "Marked as read"
    elif action == "mark_unread":
        action_script = 'set read status of aMessage to false'
        action_label = "Marked as unread"
    elif action == "flag":
        action_script = 'set flagged status of aMessage to true'
        action_label = "Flagged"
    elif action == "unflag":
        action_script = 'set flagged status of aMessage to false'
        action_label = "Unflagged"
    else:
        return f"Error: Invalid action '{action}'. Use: mark_read, mark_unread, flag, unflag"

    mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "targetMailbox")
    
    script = f'''
    tell application "Mail"
        set outputText to "UPDATING EMAIL STATUS: {action_label}" & return & return
        set updateCount to 0

        try
            set targetAccount to account "{account}"
            -- Get mailbox with dynamic inbox discovery
            try
                {mailbox_discovery}
            on error errMsg
                error "Mailbox not found: {mailbox}. " & errMsg
            end try

            set mailboxMessages to every message of targetMailbox

            repeat with aMessage in mailboxMessages
                if updateCount >= {max_updates} then exit repeat

                try
                    set messageSubject to subject of aMessage
                    set messageSender to sender of aMessage
                    set messageDate to date received of aMessage

                    -- Apply filter conditions
                    if {condition_str} then
                        {action_script}

                        set outputText to outputText & "✓ {action_label}: " & messageSubject & return
                        set outputText to outputText & "   From: " & messageSender & return
                        set outputText to outputText & "   Date: " & (messageDate as string) & return & return

                        set updateCount to updateCount + 1
                    end if
                end try
            end repeat

            set outputText to outputText & "========================================" & return
            set outputText to outputText & "TOTAL UPDATED: " & updateCount & " email(s)" & return
            set outputText to outputText & "========================================" & return

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def manage_trash(
    account: str,
    action: str,
    subject_keyword: Optional[str] = None,
    sender: Optional[str] = None,
    mailbox: str = "INBOX",
    max_deletes: int = 5,
    confirm: bool = False
) -> str:
    """
    Manage trash operations - delete emails or empty trash.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        action: Action to perform: "move_to_trash", "delete_permanent", "empty_trash"
        subject_keyword: Optional keyword to filter emails (not used for empty_trash)
        sender: Optional sender to filter emails (not used for empty_trash)
        mailbox: Source mailbox (default: "INBOX", not used for empty_trash or delete_permanent)
        max_deletes: Maximum number of emails to delete (safety limit, default: 5)
        confirm: Required for dangerous actions (empty_trash, delete_permanent).
                 If False (default), shows preview without executing.
                 If True, actually performs the deletion.
                 Note: move_to_trash does NOT require confirm (safe, reversible action).

    Returns:
        Confirmation message with details of deleted emails (preview or executed)
    """

    if action == "empty_trash":
        # empty_trash requires confirmation
        if confirm:
            delete_command = '''
                repeat with aMessage in trashMessages
                    delete aMessage
                end repeat
            '''
            status_message = "✓ Emptied trash for account: {account}"
        else:
            delete_command = '-- deletion skipped (dry run - set confirm=True to execute)'
            status_message = "📋 PREVIEW - Would empty trash for account: {account} (set confirm=True to execute)"

        script = f'''
        tell application "Mail"
            set outputText to "EMPTYING TRASH" & return & return

            try
                set targetAccount to account "{account}"
                set trashMailbox to mailbox "Trash" of targetAccount
                set trashMessages to every message of trashMailbox
                set messageCount to count of trashMessages

                -- Delete all messages in trash (or preview)
                {delete_command}

                set outputText to outputText & "{status_message}" & return
                set outputText to outputText & "   Messages in trash: " & messageCount & return

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''
    elif action == "delete_permanent":
        # delete_permanent requires confirmation
        # Build search condition
        conditions = []
        if subject_keyword:
            conditions.append(f'messageSubject contains "{subject_keyword}"')
        if sender:
            conditions.append(f'messageSender contains "{sender}"')

        condition_str = ' and '.join(conditions) if conditions else 'true'

        if confirm:
            delete_command = 'delete aMessage'
            status_message = "✓ Permanently deleted"
            header_message = "PERMANENTLY DELETING EMAILS"
        else:
            delete_command = '-- delete aMessage (dry run - set confirm=True to execute)'
            status_message = "📋 Would permanently delete"
            header_message = "PREVIEW - PERMANENT DELETION (set confirm=True to execute)"

        script = f'''
        tell application "Mail"
            set outputText to "{header_message}" & return & return
            set deleteCount to 0

            try
                set targetAccount to account "{account}"
                set trashMailbox to mailbox "Trash" of targetAccount
                set trashMessages to every message of trashMailbox

                repeat with aMessage in trashMessages
                    if deleteCount >= {max_deletes} then exit repeat

                    try
                        set messageSubject to subject of aMessage
                        set messageSender to sender of aMessage

                        -- Apply filter conditions
                        if {condition_str} then
                            set outputText to outputText & "{status_message}: " & messageSubject & return
                            set outputText to outputText & "   From: " & messageSender & return & return

                            {delete_command}
                            set deleteCount to deleteCount + 1
                        end if
                    end try
                end repeat

                set outputText to outputText & "========================================" & return
                set outputText to outputText & "TOTAL: " & deleteCount & " email(s)" & return
                set outputText to outputText & "========================================" & return

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''
    else:  # move_to_trash - safe action, no confirm needed
        # Build search condition
        conditions = []
        if subject_keyword:
            conditions.append(f'messageSubject contains "{subject_keyword}"')
        if sender:
            conditions.append(f'messageSender contains "{sender}"')

        condition_str = ' and '.join(conditions) if conditions else 'true'

        mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "sourceMailbox")
        
        script = f'''
        tell application "Mail"
            set outputText to "MOVING EMAILS TO TRASH" & return & return
            set deleteCount to 0

            try
                set targetAccount to account "{account}"
                -- Get source mailbox with dynamic inbox discovery
                try
                    {mailbox_discovery}
                on error errMsg
                    error "Mailbox not found: {mailbox}. " & errMsg
                end try

                -- Get trash mailbox
                set trashMailbox to mailbox "Trash" of targetAccount
                set sourceMessages to every message of sourceMailbox

                repeat with aMessage in sourceMessages
                    if deleteCount >= {max_deletes} then exit repeat

                    try
                        set messageSubject to subject of aMessage
                        set messageSender to sender of aMessage
                        set messageDate to date received of aMessage

                        -- Apply filter conditions
                        if {condition_str} then
                            -- Move to trash (safe, reversible action)
                            move aMessage to trashMailbox

                            set outputText to outputText & "✓ Moved to trash: " & messageSubject & return
                            set outputText to outputText & "   From: " & messageSender & return
                            set outputText to outputText & "   Date: " & (messageDate as string) & return & return

                            set deleteCount to deleteCount + 1
                        end if
                    end try
                end repeat

                set outputText to outputText & "========================================" & return
                set outputText to outputText & "TOTAL MOVED TO TRASH: " & deleteCount & " email(s)" & return
                set outputText to outputText & "========================================" & return

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def forward_email(
    account: str,
    subject_keyword: str,
    to: str,
    message: Optional[str] = None,
    mailbox: str = "INBOX",
    confirm: bool = False
) -> str:
    """
    Forward an email to one or more recipients.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        subject_keyword: Keyword to search for in email subjects
        to: Recipient email address(es), comma-separated for multiple
        message: Optional message to add before forwarded content
        mailbox: Mailbox to search in (default: "INBOX")
        confirm: If False (default), shows preview without sending.
                 If True, actually sends the forwarded email.

    Returns:
        Confirmation message with details of forwarded email (preview or sent)
    """

    escaped_message = message.replace('"', '\\"') if message else ""

    # Send command based on confirm flag
    if confirm:
        send_command = 'send forwardMessage'
        status_message = "✓ Email forwarded successfully!"
    else:
        send_command = '-- send forwardMessage (dry run - set confirm=True to send)'
        status_message = "📋 PREVIEW - Forward prepared but NOT sent (set confirm=True to send)"

    mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "targetMailbox")
    
    script = f'''
    tell application "Mail"
        set outputText to "FORWARDING EMAIL" & return & return

        try
            set targetAccount to account "{account}"
            -- Get mailbox with dynamic inbox discovery
            try
                {mailbox_discovery}
            on error errMsg
                error "Mailbox not found: {mailbox}. " & errMsg
            end try

            set mailboxMessages to every message of targetMailbox
            set foundMessage to missing value

            -- Find the first matching message
            repeat with aMessage in mailboxMessages
                try
                    set messageSubject to subject of aMessage

                    if messageSubject contains "{subject_keyword}" then
                        set foundMessage to aMessage
                        exit repeat
                    end if
                end try
            end repeat

            if foundMessage is not missing value then
                set messageSubject to subject of foundMessage
                set messageSender to sender of foundMessage
                set messageDate to date received of foundMessage

                -- Create forward
                set forwardMessage to forward foundMessage with opening window

                -- Set sender account
                set sender of forwardMessage to targetAccount

                -- Add recipients
                make new to recipient at end of to recipients of forwardMessage with properties {{address:"{to}"}}

                -- Add optional message
                if "{escaped_message}" is not "" then
                    set content of forwardMessage to "{escaped_message}" & return & return & content of forwardMessage
                end if

                -- Send the forward (or not, based on confirm)
                {send_command}

                set outputText to outputText & "{status_message}" & return & return
                set outputText to outputText & "Original email:" & return
                set outputText to outputText & "  Subject: " & messageSubject & return
                set outputText to outputText & "  From: " & messageSender & return
                set outputText to outputText & "  Date: " & (messageDate as string) & return & return
                set outputText to outputText & "Forwarded to: {to}" & return

            else
                set outputText to outputText & "⚠ No email found matching: {subject_keyword}" & return
            end if

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def get_email_thread(
    account: str,
    subject_keyword: str,
    mailbox: str = "INBOX",
    max_messages: int = 50
) -> str:
    """
    Get an email conversation thread - all messages with the same or similar subject.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        subject_keyword: Keyword to identify the thread (e.g., "Re: Project Update")
        mailbox: Mailbox to search in (default: "INBOX", use "All" for all mailboxes)
        max_messages: Maximum number of thread messages to return (default: 50)

    Returns:
        Formatted thread view with all related messages sorted by date
    """

    # For thread detection, we'll strip common prefixes
    thread_keywords = ['Re:', 'Fwd:', 'FW:', 'RE:', 'Fw:']
    cleaned_keyword = subject_keyword
    for prefix in thread_keywords:
        cleaned_keyword = cleaned_keyword.replace(prefix, '').strip()

    if mailbox == "All":
        mailbox_script = '''
            set searchMailboxes to every mailbox of targetAccount
            set useAllMailboxes to true
        '''
    else:
        mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "searchMailbox")
        mailbox_script = f'''
            try
                {mailbox_discovery}
            on error errMsg
                error "Mailbox not found: {mailbox}. " & errMsg
            end try
            set searchMailboxes to {{searchMailbox}}
            set useAllMailboxes to false
        '''

    script = f'''
    tell application "Mail"
        set outputText to "EMAIL THREAD VIEW" & return & return
        set outputText to outputText & "Thread topic: {cleaned_keyword}" & return
        set outputText to outputText & "Account: {account}" & return & return
        set threadMessages to {{}}

        try
            set targetAccount to account "{account}"
            {mailbox_script}

            -- Collect all matching messages from all mailboxes
            repeat with currentMailbox in searchMailboxes
                set mailboxMessages to every message of currentMailbox

                repeat with aMessage in mailboxMessages
                    if (count of threadMessages) >= {max_messages} then exit repeat

                    try
                        set messageSubject to subject of aMessage

                        -- Remove common prefixes for matching
                        set cleanSubject to messageSubject
                        if cleanSubject starts with "Re: " then
                            set cleanSubject to text 5 thru -1 of cleanSubject
                        end if
                        if cleanSubject starts with "Fwd: " or cleanSubject starts with "FW: " then
                            set cleanSubject to text 6 thru -1 of cleanSubject
                        end if

                        -- Check if this message is part of the thread
                        if cleanSubject contains "{cleaned_keyword}" or messageSubject contains "{cleaned_keyword}" then
                            set end of threadMessages to aMessage
                        end if
                    end try
                end repeat
            end repeat

            -- Display thread messages
            set messageCount to count of threadMessages
            set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
            set outputText to outputText & "FOUND " & messageCount & " MESSAGE(S) IN THREAD" & return
            set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

            repeat with aMessage in threadMessages
                try
                    set messageSubject to subject of aMessage
                    set messageSender to sender of aMessage
                    set messageDate to date received of aMessage
                    set messageRead to read status of aMessage

                    if messageRead then
                        set readIndicator to "✓"
                    else
                        set readIndicator to "✉"
                    end if

                    set outputText to outputText & readIndicator & " " & messageSubject & return
                    set outputText to outputText & "   From: " & messageSender & return
                    set outputText to outputText & "   Date: " & (messageDate as string) & return

                    -- Get content preview
                    try
                        set msgContent to content of aMessage
                        set AppleScript's text item delimiters to {{return, linefeed}}
                        set contentParts to text items of msgContent
                        set AppleScript's text item delimiters to " "
                        set cleanText to contentParts as string
                        set AppleScript's text item delimiters to ""

                        if length of cleanText > 150 then
                            set contentPreview to text 1 thru 150 of cleanText & "..."
                        else
                            set contentPreview to cleanText
                        end if

                        set outputText to outputText & "   Preview: " & contentPreview & return
                    end try

                    set outputText to outputText & return
                end try
            end repeat

        on error errMsg
            return "Error: " & errMsg
        end try

        return outputText
    end tell
    '''

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def manage_drafts(
    account: str,
    action: str,
    subject: Optional[str] = None,
    to: Optional[str] = None,
    body: Optional[str] = None,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    draft_subject: Optional[str] = None,
    confirm: bool = False
) -> str:
    """
    Manage draft emails - list, create, send, or delete drafts.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        action: Action to perform: "list", "create", "send", "delete"
        subject: Email subject (required for create)
        to: Recipient email(s) for create (comma-separated)
        body: Email body (required for create)
        cc: Optional CC recipients for create
        bcc: Optional BCC recipients for create
        draft_subject: Subject keyword to find draft (required for send/delete)
        confirm: Required for send and delete actions.
                 If False (default), shows preview without executing.
                 If True, actually sends or deletes the draft.
                 Note: list and create do NOT require confirm.

    Returns:
        Formatted output based on action (preview or executed)
    """

    if action == "list":
        script = f'''
        tell application "Mail"
            set outputText to "DRAFT EMAILS - {account}" & return & return

            try
                set targetAccount to account "{account}"
                set draftsMailbox to mailbox "Drafts" of targetAccount
                set draftMessages to every message of draftsMailbox
                set draftCount to count of draftMessages

                set outputText to outputText & "Found " & draftCount & " draft(s)" & return & return

                repeat with aDraft in draftMessages
                    try
                        set draftSubject to subject of aDraft
                        set draftDate to date sent of aDraft

                        set outputText to outputText & "✉ " & draftSubject & return
                        set outputText to outputText & "   Created: " & (draftDate as string) & return & return
                    end try
                end repeat

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    elif action == "create":
        if not subject or not to or not body:
            return "Error: 'subject', 'to', and 'body' are required for creating drafts"

        escaped_subject = subject.replace('"', '\\"')
        escaped_body = body.replace('"', '\\"')

        # Build CC recipients if provided
        cc_script = ''
        if cc:
            cc_addresses = [addr.strip() for addr in cc.split(',')]
            for addr in cc_addresses:
                cc_script += f'''
                make new cc recipient at end of cc recipients of newDraft with properties {{address:"{addr}"}}
                '''

        # Build BCC recipients if provided
        bcc_script = ''
        if bcc:
            bcc_addresses = [addr.strip() for addr in bcc.split(',')]
            for addr in bcc_addresses:
                bcc_script += f'''
                make new bcc recipient at end of bcc recipients of newDraft with properties {{address:"{addr}"}}
                '''

        script = f'''
        tell application "Mail"
            set outputText to "CREATING DRAFT" & return & return

            try
                set targetAccount to account "{account}"

                -- Create new outgoing message (draft)
                set newDraft to make new outgoing message with properties {{subject:"{escaped_subject}", content:"{escaped_body}", visible:false}}

                -- Set the sender account
                set sender of newDraft to targetAccount

                -- Add recipients
                tell newDraft
                    make new to recipient at end of to recipients with properties {{address:"{to}"}}
                    {cc_script}
                    {bcc_script}
                end tell

                -- Save to drafts (don't send)
                -- The draft is automatically saved to Drafts folder

                set outputText to outputText & "✓ Draft created successfully!" & return & return
                set outputText to outputText & "Subject: {escaped_subject}" & return
                set outputText to outputText & "To: {to}" & return

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    elif action == "send":
        if not draft_subject:
            return "Error: 'draft_subject' is required for sending drafts"

        # send requires confirmation
        if confirm:
            send_command = 'send foundDraft'
            status_message = "✓ Draft sent successfully!"
            header_message = "SENDING DRAFT"
        else:
            send_command = '-- send foundDraft (dry run - set confirm=True to send)'
            status_message = "📋 PREVIEW - Draft found but NOT sent (set confirm=True to send)"
            header_message = "PREVIEW - SEND DRAFT"

        script = f'''
        tell application "Mail"
            set outputText to "{header_message}" & return & return

            try
                set targetAccount to account "{account}"
                set draftsMailbox to mailbox "Drafts" of targetAccount
                set draftMessages to every message of draftsMailbox
                set foundDraft to missing value

                -- Find the draft
                repeat with aDraft in draftMessages
                    try
                        set draftSubject to subject of aDraft

                        if draftSubject contains "{draft_subject}" then
                            set foundDraft to aDraft
                            exit repeat
                        end if
                    end try
                end repeat

                if foundDraft is not missing value then
                    set draftSubject to subject of foundDraft

                    -- Send the draft (or preview)
                    {send_command}

                    set outputText to outputText & "{status_message}" & return
                    set outputText to outputText & "Subject: " & draftSubject & return

                else
                    set outputText to outputText & "⚠ No draft found matching: {draft_subject}" & return
                end if

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    elif action == "delete":
        if not draft_subject:
            return "Error: 'draft_subject' is required for deleting drafts"

        # delete requires confirmation
        if confirm:
            delete_command = 'delete foundDraft'
            status_message = "✓ Draft deleted successfully!"
            header_message = "DELETING DRAFT"
        else:
            delete_command = '-- delete foundDraft (dry run - set confirm=True to delete)'
            status_message = "📋 PREVIEW - Draft found but NOT deleted (set confirm=True to delete)"
            header_message = "PREVIEW - DELETE DRAFT"

        script = f'''
        tell application "Mail"
            set outputText to "{header_message}" & return & return

            try
                set targetAccount to account "{account}"
                set draftsMailbox to mailbox "Drafts" of targetAccount
                set draftMessages to every message of draftsMailbox
                set foundDraft to missing value

                -- Find the draft
                repeat with aDraft in draftMessages
                    try
                        set draftSubject to subject of aDraft

                        if draftSubject contains "{draft_subject}" then
                            set foundDraft to aDraft
                            exit repeat
                        end if
                    end try
                end repeat

                if foundDraft is not missing value then
                    set draftSubject to subject of foundDraft

                    -- Delete the draft (or preview)
                    {delete_command}

                    set outputText to outputText & "{status_message}" & return
                    set outputText to outputText & "Subject: " & draftSubject & return

                else
                    set outputText to outputText & "⚠ No draft found matching: {draft_subject}" & return
                end if

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    else:
        return f"Error: Invalid action '{action}'. Use: list, create, send, delete"

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def get_statistics(
    account: str,
    scope: str = "account_overview",
    sender: Optional[str] = None,
    mailbox: Optional[str] = None,
    days_back: int = 30
) -> str:
    """
    Get comprehensive email statistics and analytics.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        scope: Analysis scope: "account_overview", "sender_stats", "mailbox_breakdown"
        sender: Specific sender for "sender_stats" scope
        mailbox: Specific mailbox for "mailbox_breakdown" scope
        days_back: Number of days to analyze (default: 30, 0 = all time)

    Returns:
        Formatted statistics report with metrics and insights
    """

    # Calculate date threshold if days_back > 0
    date_filter = ""
    if days_back > 0:
        date_filter = f'''
            set targetDate to (current date) - ({days_back} * days)
        '''
        date_check = 'and messageDate > targetDate'
    else:
        date_filter = ""
        date_check = ""

    if scope == "account_overview":
        script = f'''
        tell application "Mail"
            set outputText to "╔══════════════════════════════════════════╗" & return
            set outputText to outputText & "║      EMAIL STATISTICS - {account}       ║" & return
            set outputText to outputText & "╚══════════════════════════════════════════╝" & return & return

            {date_filter}

            try
                set targetAccount to account "{account}"
                set allMailboxes to every mailbox of targetAccount

                -- Initialize counters
                set totalEmails to 0
                set totalUnread to 0
                set totalRead to 0
                set totalFlagged to 0
                set totalWithAttachments to 0
                set senderCounts to {{}}
                set mailboxCounts to {{}}

                -- Analyze all mailboxes
                repeat with aMailbox in allMailboxes
                    set mailboxName to name of aMailbox
                    set mailboxMessages to every message of aMailbox
                    set mailboxTotal to 0

                    repeat with aMessage in mailboxMessages
                        try
                            set messageDate to date received of aMessage

                            -- Apply date filter if specified
                            if true {date_check} then
                                set totalEmails to totalEmails + 1
                                set mailboxTotal to mailboxTotal + 1

                                -- Count read/unread
                                if read status of aMessage then
                                    set totalRead to totalRead + 1
                                else
                                    set totalUnread to totalUnread + 1
                                end if

                                -- Count flagged
                                try
                                    if flagged status of aMessage then
                                        set totalFlagged to totalFlagged + 1
                                    end if
                                end try

                                -- Count attachments
                                set attachmentCount to count of mail attachments of aMessage
                                if attachmentCount > 0 then
                                    set totalWithAttachments to totalWithAttachments + 1
                                end if

                                -- Track senders (top 10)
                                set messageSender to sender of aMessage
                                set senderFound to false
                                repeat with senderPair in senderCounts
                                    if item 1 of senderPair is messageSender then
                                        set item 2 of senderPair to (item 2 of senderPair) + 1
                                        set senderFound to true
                                        exit repeat
                                    end if
                                end repeat
                                if not senderFound then
                                    set end of senderCounts to {{messageSender, 1}}
                                end if
                            end if
                        end try
                    end repeat

                    -- Store mailbox counts
                    if mailboxTotal > 0 then
                        set end of mailboxCounts to {{mailboxName, mailboxTotal}}
                    end if
                end repeat

                -- Format output
                set outputText to outputText & "📊 VOLUME METRICS" & return
                set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
                set outputText to outputText & "Total Emails: " & totalEmails & return
                set outputText to outputText & "Unread: " & totalUnread & " (" & (round ((totalUnread / totalEmails) * 100)) & "%)" & return
                set outputText to outputText & "Read: " & totalRead & " (" & (round ((totalRead / totalEmails) * 100)) & "%)" & return
                set outputText to outputText & "Flagged: " & totalFlagged & return
                set outputText to outputText & "With Attachments: " & totalWithAttachments & " (" & (round ((totalWithAttachments / totalEmails) * 100)) & "%)" & return
                set outputText to outputText & return

                -- Top senders (show top 5)
                set outputText to outputText & "👥 TOP SENDERS" & return
                set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
                set topCount to 0
                repeat with senderPair in senderCounts
                    set topCount to topCount + 1
                    if topCount > 5 then exit repeat
                    set outputText to outputText & item 1 of senderPair & ": " & item 2 of senderPair & " emails" & return
                end repeat
                set outputText to outputText & return

                -- Mailbox distribution (show top 5)
                set outputText to outputText & "📁 MAILBOX DISTRIBUTION" & return
                set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
                set topCount to 0
                repeat with mailboxPair in mailboxCounts
                    set topCount to topCount + 1
                    if topCount > 5 then exit repeat
                    set mailboxPercent to round ((item 2 of mailboxPair / totalEmails) * 100)
                    set outputText to outputText & item 1 of mailboxPair & ": " & item 2 of mailboxPair & " (" & mailboxPercent & "%)" & return
                end repeat

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    elif scope == "sender_stats":
        if not sender:
            return "Error: 'sender' parameter required for sender_stats scope"

        script = f'''
        tell application "Mail"
            set outputText to "SENDER STATISTICS" & return & return
            set outputText to outputText & "Sender: {sender}" & return
            set outputText to outputText & "Account: {account}" & return & return

            {date_filter}

            try
                set targetAccount to account "{account}"
                set allMailboxes to every mailbox of targetAccount

                set totalFromSender to 0
                set unreadFromSender to 0
                set withAttachments to 0

                repeat with aMailbox in allMailboxes
                    set mailboxMessages to every message of aMailbox

                    repeat with aMessage in mailboxMessages
                        try
                            set messageSender to sender of aMessage
                            set messageDate to date received of aMessage

                            if messageSender contains "{sender}" {date_check} then
                                set totalFromSender to totalFromSender + 1

                                if not (read status of aMessage) then
                                    set unreadFromSender to unreadFromSender + 1
                                end if

                                if (count of mail attachments of aMessage) > 0 then
                                    set withAttachments to withAttachments + 1
                                end if
                            end if
                        end try
                    end repeat
                end repeat

                set outputText to outputText & "Total emails: " & totalFromSender & return
                set outputText to outputText & "Unread: " & unreadFromSender & return
                set outputText to outputText & "With attachments: " & withAttachments & return

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    elif scope == "mailbox_breakdown":
        mailbox_param = mailbox if mailbox else "INBOX"
        mailbox_discovery = get_mailbox_applescript_snippet(mailbox_param, "targetAccount", "targetMailbox")

        script = f'''
        tell application "Mail"
            set outputText to "MAILBOX STATISTICS" & return & return
            set outputText to outputText & "Mailbox: {mailbox_param}" & return
            set outputText to outputText & "Account: {account}" & return & return

            try
                set targetAccount to account "{account}"
                try
                    {mailbox_discovery}
                on error errMsg
                    error "Mailbox not found: {mailbox_param}. " & errMsg
                end try

                set mailboxMessages to every message of targetMailbox
                set totalMessages to count of mailboxMessages
                set unreadMessages to unread count of targetMailbox

                set outputText to outputText & "Total messages: " & totalMessages & return
                set outputText to outputText & "Unread: " & unreadMessages & return
                set outputText to outputText & "Read: " & (totalMessages - unreadMessages) & return

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    else:
        return f"Error: Invalid scope '{scope}'. Use: account_overview, sender_stats, mailbox_breakdown"

    result = await run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
async def export_emails(
    account: str,
    scope: str,
    subject_keyword: Optional[str] = None,
    mailbox: str = "INBOX",
    save_directory: str = "~/Desktop",
    format: str = "txt"
) -> str:
    """
    Export emails to files for backup or analysis.

    Args:
        account: Account name (e.g., "Gmail", "Work")
        scope: Export scope: "single_email" (requires subject_keyword) or "entire_mailbox"
        subject_keyword: Keyword to find email (required for single_email)
        mailbox: Mailbox to export from (default: "INBOX")
        save_directory: Directory to save exports (default: "~/Desktop")
        format: Export format: "txt", "html" (default: "txt")

    Returns:
        Confirmation message with export location
    """

    # Expand home directory
    import os
    save_dir = os.path.expanduser(save_directory)

    if scope == "single_email":
        if not subject_keyword:
            return "Error: 'subject_keyword' required for single_email scope"

        mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "targetMailbox")
        
        script = f'''
        tell application "Mail"
            set outputText to "EXPORTING EMAIL" & return & return

            try
                set targetAccount to account "{account}"
                -- Get mailbox with dynamic inbox discovery
                try
                    {mailbox_discovery}
                on error errMsg
                    error "Mailbox not found: {mailbox}. " & errMsg
                end try

                set mailboxMessages to every message of targetMailbox
                set foundMessage to missing value

                -- Find the email
                repeat with aMessage in mailboxMessages
                    try
                        set messageSubject to subject of aMessage

                        if messageSubject contains "{subject_keyword}" then
                            set foundMessage to aMessage
                            exit repeat
                        end if
                    end try
                end repeat

                if foundMessage is not missing value then
                    set messageSubject to subject of foundMessage
                    set messageSender to sender of foundMessage
                    set messageDate to date received of foundMessage
                    set messageContent to content of foundMessage

                    -- Create safe filename
                    set safeSubject to messageSubject
                    set AppleScript's text item delimiters to "/"
                    set safeSubjectParts to text items of safeSubject
                    set AppleScript's text item delimiters to "-"
                    set safeSubject to safeSubjectParts as string
                    set AppleScript's text item delimiters to ""

                    set fileName to safeSubject & ".{format}"
                    set filePath to "{save_dir}/" & fileName

                    -- Prepare export content
                    if "{format}" is "txt" then
                        set exportContent to "Subject: " & messageSubject & return
                        set exportContent to exportContent & "From: " & messageSender & return
                        set exportContent to exportContent & "Date: " & (messageDate as string) & return & return
                        set exportContent to exportContent & messageContent
                    else if "{format}" is "html" then
                        set exportContent to "<html><body>"
                        set exportContent to exportContent & "<h2>" & messageSubject & "</h2>"
                        set exportContent to exportContent & "<p><strong>From:</strong> " & messageSender & "</p>"
                        set exportContent to exportContent & "<p><strong>Date:</strong> " & (messageDate as string) & "</p>"
                        set exportContent to exportContent & "<hr>" & messageContent
                        set exportContent to exportContent & "</body></html>"
                    end if

                    -- Write to file
                    set fileRef to open for access POSIX file filePath with write permission
                    set eof of fileRef to 0
                    write exportContent to fileRef as «class utf8»
                    close access fileRef

                    set outputText to outputText & "✓ Email exported successfully!" & return & return
                    set outputText to outputText & "Subject: " & messageSubject & return
                    set outputText to outputText & "Saved to: " & filePath & return

                else
                    set outputText to outputText & "⚠ No email found matching: {subject_keyword}" & return
                end if

            on error errMsg
                try
                    close access file filePath
                end try
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    elif scope == "entire_mailbox":
        mailbox_discovery = get_mailbox_applescript_snippet(mailbox, "targetAccount", "targetMailbox")
        
        script = f'''
        tell application "Mail"
            set outputText to "EXPORTING MAILBOX" & return & return

            try
                set targetAccount to account "{account}"
                -- Get mailbox with dynamic inbox discovery
                try
                    {mailbox_discovery}
                on error errMsg
                    error "Mailbox not found: {mailbox}. " & errMsg
                end try

                set mailboxMessages to every message of targetMailbox
                set messageCount to count of mailboxMessages
                set exportCount to 0

                -- Create export directory
                set exportDir to "{save_dir}/{mailbox}_export"
                do shell script "mkdir -p " & quoted form of exportDir

                repeat with aMessage in mailboxMessages
                    try
                        set messageSubject to subject of aMessage
                        set messageSender to sender of aMessage
                        set messageDate to date received of aMessage
                        set messageContent to content of aMessage

                        -- Create safe filename with index
                        set exportCount to exportCount + 1
                        set fileName to exportCount & "_" & messageSubject & ".{format}"

                        -- Remove unsafe characters
                        set AppleScript's text item delimiters to "/"
                        set fileNameParts to text items of fileName
                        set AppleScript's text item delimiters to "-"
                        set fileName to fileNameParts as string
                        set AppleScript's text item delimiters to ""

                        set filePath to exportDir & "/" & fileName

                        -- Prepare export content
                        if "{format}" is "txt" then
                            set exportContent to "Subject: " & messageSubject & return
                            set exportContent to exportContent & "From: " & messageSender & return
                            set exportContent to exportContent & "Date: " & (messageDate as string) & return & return
                            set exportContent to exportContent & messageContent
                        else if "{format}" is "html" then
                            set exportContent to "<html><body>"
                            set exportContent to exportContent & "<h2>" & messageSubject & "</h2>"
                            set exportContent to exportContent & "<p><strong>From:</strong> " & messageSender & "</p>"
                            set exportContent to exportContent & "<p><strong>Date:</strong> " & (messageDate as string) & "</p>"
                            set exportContent to exportContent & "<hr>" & messageContent
                            set exportContent to exportContent & "</body></html>"
                        end if

                        -- Write to file
                        set fileRef to open for access POSIX file filePath with write permission
                        set eof of fileRef to 0
                        write exportContent to fileRef as «class utf8»
                        close access fileRef

                    on error
                        -- Continue with next email if one fails
                    end try
                end repeat

                set outputText to outputText & "✓ Mailbox exported successfully!" & return & return
                set outputText to outputText & "Mailbox: {mailbox}" & return
                set outputText to outputText & "Total emails: " & messageCount & return
                set outputText to outputText & "Exported: " & exportCount & return
                set outputText to outputText & "Location: " & exportDir & return

            on error errMsg
                return "Error: " & errMsg
            end try

            return outputText
        end tell
        '''

    else:
        return f"Error: Invalid scope '{scope}'. Use: single_email, entire_mailbox"

    result = await run_applescript(script)
    return result


if __name__ == "__main__":
    # Run the MCP server
    mcp.run()
