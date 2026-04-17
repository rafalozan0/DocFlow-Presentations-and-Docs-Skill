"""DocFlow - Presentations and Docs Skill.

Agent-friendly office automation helpers for Word, Excel, PDF, and PowerPoint.
"""

__version__ = "1.1.0"
__author__ = "Tao Jin (shynloc) [original], Rafael Lozano [modifications] & contributors"

from . import template_registry
from .core import OfficeSuite

__all__ = ["OfficeSuite", "template_registry"]
