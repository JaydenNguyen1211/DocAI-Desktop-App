from dataclasses import dataclass


@dataclass
class AttachedFile:
    path: str
    name: str
    file_type: str        # 'word' | 'excel'
    claude_content: str
