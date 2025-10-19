# Temporary compatibility shim for Python 3.13+
# Streamlit <1.39 still imports imghdr, which was removed from stdlib.

import io

def what(file, h=None):
    """Always return 'png' if bytes start with PNG header, otherwise None."""
    if h is None:
        if hasattr(file, 'read'):
            pos = file.tell()
            h = file.read(32)
            file.seek(pos)
        else:
            return None
    if isinstance(h, (bytes, bytearray)) and h.startswith(b'\211PNG\r\n\032\n'):
        return "png"
    return None
