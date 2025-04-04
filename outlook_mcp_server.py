import sys
import traceback
import datetime
import os
import json
import logging
import win32com.client
from typing import List, Optional, Dict, Any
from mcp.server.fastmcp import FastMCP, Context

# Configure logging to both file and stderr
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('outlook_mcp_server.log'),
        logging.StreamHandler(sys.stderr)
    ]
)

# Global error handler to log all unhandled exceptions
def global_exception_handler(exctype, value, tb):
    logging.error("Uncaught exception:", exc_info=(exctype, value, tb))
    # Print to stderr to ensure visibility
    traceback.print_exception(exctype, value, tb, file=sys.stderr)

sys.excepthook = global_exception_handler

# Initialize FastMCP server with more robust configuration
try:
    mcp = FastMCP("outlook-assistant", 
                  error_handling=True,  # Enable built-in error handling
                  log_level="DEBUG")  # Use string instead of integer
except Exception as e:
    logging.error(f"Failed to initialize FastMCP: {e}")
    sys.exit(1)

# Constants
MAX_DAYS = 30
# Email cache for storing retrieved emails by number
email_cache = {}

# Helper functions
def safe_connect_to_outlook():
    """
    Safely connect to Outlook with comprehensive error handling
    
    Returns:
        tuple: (outlook application, namespace)
    
    Raises:
        Exception with detailed error information
    """
    try:
        # Attempt to import required libraries
        import win32com.client
        
        # Attempt to dispatch Outlook application
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as dispatch_error:
            logging.error(f"Failed to dispatch Outlook: {dispatch_error}")
            raise Exception(f"Failed to connect to Outlook: {dispatch_error}")
        
        # Verify basic connectivity
        try:
            inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
            logging.info(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
        except Exception as connectivity_error:
            logging.error(f"Outlook connection verified, but inbox access failed: {connectivity_error}")
            raise Exception(f"Inbox access failed: {connectivity_error}")
        
        return outlook, namespace
    
    except ImportError:
        logging.error("win32com.client library not found. Please install pywin32.")
        raise Exception("win32com.client library not found. Please install pywin32.")
    except Exception as e:
        logging.error(f"Outlook connection failed: {e}")
        raise

# Helper function to get folder by name
def get_folder_by_name(namespace, folder_name: Optional[str] = None):
    """
    Get a specific Outlook folder by name. If no name is provided, return the default inbox.
    
    Args:
        namespace: Outlook namespace
        folder_name: Name of the folder to retrieve
    
    Returns:
        Outlook folder object
    """
    try:
        if not folder_name:
            return namespace.GetDefaultFolder(6)  # 6 is inbox
        
        # First check inbox subfolder
        inbox = namespace.GetDefaultFolder(6)
        
        # Check inbox subfolders first (most common)
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
                
        # Then check all folders at root level
        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
            
            # Also check subfolders
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
                    
        # If not found
        raise Exception(f"Folder '{folder_name}' not found")
    except Exception as e:
        logging.error(f"Failed to access folder {folder_name}: {e}")
        raise

# Folder listing tool
@mcp.tool()
def outlook_list_folders() -> str:
    """
    List all available mail folders in Outlook
    
    Returns:
        A list of available mail folders
    """
    try:
        # Connect to Outlook
        _, namespace = safe_connect_to_outlook()
        
        result = "Available mail folders:\n\n"
        
        # List all root folders and their subfolders
        for folder in namespace.Folders:
            result += f"- {folder.Name}\n"
            
            # List subfolders
            for subfolder in folder.Folders:
                result += f"  - {subfolder.Name}\n"
                
                # List subfolders (one more level)
                try:
                    for subsubfolder in subfolder.Folders:
                        result += f"    - {subsubfolder.Name}\n"
                except:
                    pass
        
        return result
    except Exception as e:
        logging.error(f"Failed to list folders: {e}")
        return f"Error listing mail folders: {e}"

# List recent emails tool
@mcp.tool()
def outlook_list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> Dict:
    """
    List recent emails from a specified folder
    
    Args:
        days: Number of days to look back (default 7)
        folder_name: Optional folder name to search (default is Inbox)
    
    Returns:
        Dictionary with email details
    """
    try:
        # Validate days parameter
        if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
            raise ValueError(f"Days must be between 1 and {MAX_DAYS}")
        
        # Connect to Outlook
        _, namespace = safe_connect_to_outlook()
        
        # Get the appropriate folder
        folder = get_folder_by_name(namespace, folder_name)
        
        # Clear previous email cache
        global email_cache
        email_cache.clear()
        
        # Calculate the date threshold
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        # Process emails
        emails_list = []
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        for i, item in enumerate(folder_items, 1):
            # Check if the email is within the time threshold
            if hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                received_time = item.ReceivedTime.replace(tzinfo=None)
                
                if received_time < threshold_date:
                    break
                
                # Format email details
                email_data = {
                    "id": str(i),
                    "subject": item.Subject or "",
                    "sender": item.SenderName or "",
                    "sender_email": item.SenderEmailAddress or "",
                    "received_time": received_time.strftime("%Y-%m-%d %H:%M:%S"),
                    "read_status": "Read" if not item.UnRead else "Unread",
                    "has_attachments": item.Attachments.Count > 0
                }
                
                # Cache the email
                email_cache[i] = item
                emails_list.append(email_data)
        
        return {
            "folder": folder.Name,
            "total_emails": len(emails_list),
            "emails": emails_list
        }
    
    except Exception as e:
        logging.error(f"Error retrieving emails: {e}")
        return {"error": str(e)}

# Search emails tool
@mcp.tool()
def outlook_search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None) -> Dict:
    """
    Search emails by keyword within a time period
    
    Args:
        search_term: Keyword to search for
        days: Number of days to look back (default 7)
        folder_name: Optional folder name to search (default is Inbox)
    
    Returns:
        Dictionary with matching email details
    """
    try:
        # Validate parameters
        if not search_term:
            raise ValueError("Search term cannot be empty")
        
        if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
            raise ValueError(f"Days must be between 1 and {MAX_DAYS}")
        
        # Connect to Outlook
        _, namespace = safe_connect_to_outlook()
        
        # Get the appropriate folder
        folder = get_folder_by_name(namespace, folder_name)
        
        # Clear previous email cache
        global email_cache
        email_cache.clear()
        
        # Calculate the date threshold
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        # Process emails
        emails_list = []
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        # Normalize search term
        search_term_lower = search_term.lower()
        
        for i, item in enumerate(folder_items, 1):
            # Check if the email is within the time threshold
            if hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                received_time = item.ReceivedTime.replace(tzinfo=None)
                
                if received_time < threshold_date:
                    break
                
                # Search in subject, sender, and body
                try:
                    # Check for match in subject, sender name, or body
                    if (search_term_lower in (item.Subject or "").lower() or 
                        search_term_lower in (item.SenderName or "").lower() or 
                        search_term_lower in (item.Body or "").lower()):
                        
                        # Format email details
                        email_data = {
                            "id": str(i),
                            "subject": item.Subject or "",
                            "sender": item.SenderName or "",
                            "sender_email": item.SenderEmailAddress or "",
                            "received_time": received_time.strftime("%Y-%m-%d %H:%M:%S"),
                            "read_status": "Read" if not item.UnRead else "Unread",
                            "has_attachments": item.Attachments.Count > 0
                        }
                        
                        # Cache the email
                        email_cache[i] = item
                        emails_list.append(email_data)
                except Exception as email_error:
                    logging.warning(f"Error processing email: {email_error}")
        
        return {
            "folder": folder.Name,
            "search_term": search_term,
            "total_emails": len(emails_list),
            "emails": emails_list
        }
    
    except Exception as e:
        logging.error(f"Error searching emails: {e}")
        return {"error": str(e)}

# Get email by number tool
@mcp.tool()
def outlook_get_email_by_number(email_number: int) -> Dict:
    """
    Retrieve detailed information for a specific email by its number
    
    Args:
        email_number: Number of the email from previous listing
    
    Returns:
        Dictionary with detailed email information
    """
    try:
        # Retrieve the email from the cache
        if not email_cache:
            return {"error": "No emails have been listed recently. Use list_recent_emails or search_emails first."}
        
        if email_number not in email_cache:
            return {"error": f"Email #{email_number} not found in the current listing."}
        
        # Get the email item
        email_item = email_cache[email_number]
        
        # Extract detailed information
        email_details = {
            "subject": email_item.Subject or "",
            "sender": email_item.SenderName or "",
            "sender_email": email_item.SenderEmailAddress or "",
            "received_time": email_item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if email_item.ReceivedTime else "",
            "body": email_item.Body or "",
            "read_status": "Read" if not email_item.UnRead else "Unread",
            "importance": email_item.Importance if hasattr(email_item, 'Importance') else 1,
            "has_attachments": email_item.Attachments.Count > 0,
            "attachments": []
        }
        
        # Add attachment details if present
        if email_details["has_attachments"]:
            for i in range(1, email_item.Attachments.Count + 1):
                attachment = email_item.Attachments(i)
                email_details["attachments"].append({
                    "filename": attachment.FileName,
                    "size": attachment.Size
                })
        
        return email_details
    
    except Exception as e:
        logging.error(f"Error retrieving email details: {e}")
        return {"error": str(e)}

# Run the server
def main():
    try:
        logging.info("Starting Outlook MCP Server in READ-ONLY MODE...")
        
        # Pre-flight checks
        try:
            # Test Outlook connection
            safe_connect_to_outlook()
        except Exception as conn_error:
            logging.error(f"Pre-flight Outlook connection check failed: {conn_error}")
            sys.exit(1)
        
        # Run the MCP server
        logging.info("Starting MCP server. Press Ctrl+C to stop.")
        mcp.run()
    
    except KeyboardInterrupt:
        logging.info("Server stopped by user.")
    except Exception as e:
        logging.error(f"Unhandled server error: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
