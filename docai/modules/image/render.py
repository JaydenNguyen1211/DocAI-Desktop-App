import os
import tempfile

from ...imkit import EagerPageSource

from ...logging_config import get_logger, log_call

logger = get_logger(__name__)


@log_call
def image_page_source(path: str) -> EagerPageSource:
    from PIL import Image

    tmp_dir = tempfile.mkdtemp(prefix="docai_pages_")
    out = os.path.join(tmp_dir, "page_000.png")
    Image.open(path).convert("RGB").save(out, "PNG")
    return EagerPageSource([out], tmp_dir)
