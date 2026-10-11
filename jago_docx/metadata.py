"""
Metadata manager for Jago Docx.
Ensures the current OS/environment username is automatically assigned as the document author,
and strictly removes/prevents default strings like 'python-docx' from appearing in document properties.
"""

import os
import getpass
from typing import Optional


def get_default_username() -> str:
    """
    Retrieve the current logged-in OS username.
    Checks environment variables across Windows (USERNAME) and POSIX (USER),
    falling back to getpass.getuser() or a clean default.
    """
    username = (
        os.environ.get("USERNAME")
        or os.environ.get("USER")
        or os.environ.get("LOGNAME")
    )
    if not username:
        try:
            username = getpass.getuser()
        except Exception:
            username = "Author"
    
    # Strip any invalid characters or whitespace
    username = username.strip() if username else "Author"
    return username


def apply_author_metadata(doc, author: Optional[str] = None, title: Optional[str] = None, subject: Optional[str] = None):
    """
    Set core document properties with author information.
    Ensures 'python-docx' is never used as author/creator.
    """
    effective_author = author.strip() if author and author.strip() else get_default_username()
    
    core_props = doc.core_properties
    core_props.author = effective_author
    core_props.last_modified_by = effective_author
    
    # In some python-docx versions or templates, comments or other fields might be set.
    # We ensure they are clean.
    if title:
        core_props.title = title
    if subject:
        core_props.subject = subject
    
    # Scrub 'python-docx' if present in any string property
    for attr in ("author", "last_modified_by", "comments", "category", "keywords"):
        val = getattr(core_props, attr, None)
        if val and isinstance(val, str) and "python-docx" in val.lower():
            setattr(core_props, attr, val.replace("python-docx", effective_author))
            
    return effective_author
