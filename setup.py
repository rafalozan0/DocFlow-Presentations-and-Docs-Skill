from pathlib import Path
import re

from setuptools import find_packages, setup

ROOT = Path(__file__).parent
INIT_FILE = ROOT / "src" / "office_suite" / "__init__.py"


def _read_metadata() -> tuple[str, str]:
    content = INIT_FILE.read_text(encoding="utf-8")
    version_match = re.search(r"__version__\s*=\s*\"([^\"]+)\"", content)
    author_match = re.search(r"__author__\s*=\s*\"([^\"]+)\"", content)

    if not version_match or not author_match:
        raise RuntimeError("Could not parse __version__ or __author__ from src/office_suite/__init__.py")

    return version_match.group(1), author_match.group(1)


__version__, __author__ = _read_metadata()

long_description = (ROOT / "README.md").read_text(encoding="utf-8")
requirements = [
    line.strip()
    for line in (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines()
    if line.strip() and not line.startswith("#")
]

setup(
    name="docflow-presentations-and-docs-skill",
    version=__version__,
    author=__author__,
    author_email="mail@jintao.uk",
    description="Agent-friendly Python office automation toolkit for Word/Excel/PDF/PPTX.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/rafalozan0/DocFlow-Presentations-and-Docs-Skill",
    project_urls={
        "Bug Tracker": "https://github.com/rafalozan0/DocFlow-Presentations-and-Docs-Skill/issues",
        "Source": "https://github.com/rafalozan0/DocFlow-Presentations-and-Docs-Skill",
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Natural Language :: English",
        "Natural Language :: Spanish",
    ],
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "full": [
            "Jinja2>=3.1.0",
            "click>=8.1.0",
            "rich>=13.0.0",
            "python-magic>=0.4.27",
            "python-dotenv>=1.0.0",
        ],
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
        ],
    },
    keywords=[
        "office",
        "excel",
        "word",
        "powerpoint",
        "pdf",
        "docx",
        "xlsx",
        "pptx",
        "automation",
        "report-generator",
        "hermes",
        "openclaw",
        "ai-agent",
    ],
    include_package_data=True,
)
