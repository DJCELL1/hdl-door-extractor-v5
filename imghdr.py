# Compatibility shim for Python 3.13+
# Streamlit <1.39 still imports 'imghdr', which was removed.

import io

def what(file, h=None):
    """Return 'png' for PNG headers, otherwise None."""
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

