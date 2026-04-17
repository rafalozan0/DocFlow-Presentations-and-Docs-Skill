import os
from typing import Any, Dict, List, Optional

from . import docx, html_slides, pdf, pptx, template_registry, utils, xlsx


class OfficeSuite:
    """Unified API for Office document operations.

    Supported types:
    - word/docx
    - excel/xlsx
    - pdf
    - pptx/powerpoint
    """

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or {}

    @staticmethod
    def _normalize_output_kw(kwargs: Dict[str, Any]) -> Dict[str, Any]:
        """Allow both `output` and `output_path` for compatibility."""
        normalized = dict(kwargs)
        if "output" in normalized and "output_path" not in normalized:
            normalized["output_path"] = normalized.pop("output")
        return normalized

    def create(self, doc_type: str, **kwargs) -> Dict[str, Any]:
        """Create Office documents.

        For PPTX generation, you can enforce preflight preferences with:
        - require_preflight=True
        and providing:
        - theme, chart_mode, use_emojis, tone
        """
        doc_type = doc_type.lower().strip()
        handlers = {
            "word": docx.create_word,
            "docx": docx.create_word,
            "excel": xlsx.create_excel,
            "xlsx": xlsx.create_excel,
            "pdf": pdf.create_pdf,
            "pptx": pptx.create_pptx,
            "powerpoint": pptx.create_pptx,
            "pptx_html": html_slides.build_pptx_from_html_theme,
            "html_pptx": html_slides.build_pptx_from_html_theme,
            "slide_template": self._create_single_slide_from_template,
            "slide_deck_templates": self._create_pptx_from_templates,
            "doc_template": self._create_document_from_template,
        }

        if doc_type not in handlers:
            return {"success": False, "error": f"Unsupported document type: {doc_type}"}

        try:
            normalized_kwargs = self._normalize_output_kw(kwargs)
            result = handlers[doc_type](**normalized_kwargs)
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": f"Create operation failed: {e}"}

    def convert(
        self,
        input_path: str,
        to: str,
        output_path: Optional[str] = None,
        **kwargs,
    ) -> Dict[str, Any]:
        """Convert a document using LibreOffice headless mode."""
        if not os.path.exists(input_path):
            return {"success": False, "error": f"Input file not found: {input_path}"}

        if not output_path and "output" in kwargs:
            output_path = kwargs.pop("output")

        try:
            if not output_path:
                input_dir = os.path.dirname(input_path)
                input_name = os.path.splitext(os.path.basename(input_path))[0]
                output_path = os.path.join(input_dir, f"{input_name}.{to.lower()}")

            target_ext = to.lower().lstrip(".")
            result = utils.convert_with_libreoffice(input_path, target_ext, output_path, **kwargs)
            return {"success": True, "output_path": output_path, **result}
        except Exception as e:
            return {"success": False, "error": f"Conversion failed: {e}"}

    def batch_convert(
        self,
        source_dir: str,
        target_dir: str,
        source_format: str,
        target_format: str,
        **kwargs,
    ) -> Dict[str, Dict[str, Any]]:
        """Batch convert documents from source_dir to target_dir."""
        if not os.path.exists(source_dir):
            return {"error": f"Source directory not found: {source_dir}"}

        os.makedirs(target_dir, exist_ok=True)
        results: Dict[str, Dict[str, Any]] = {}

        for filename in os.listdir(source_dir):
            if filename.lower().endswith(f".{source_format.lower()}"):
                input_path = os.path.join(source_dir, filename)
                output_name = f"{os.path.splitext(filename)[0]}.{target_format.lower()}"
                output_path = os.path.join(target_dir, output_name)
                results[filename] = self.convert(input_path, target_format, output_path, **kwargs)

        return results

    def add_watermark(
        self,
        input_path: str,
        watermark_text: str,
        output_path: Optional[str] = None,
        **kwargs,
    ) -> Dict[str, Any]:
        """Add watermark to PDF or DOCX documents."""
        if not os.path.exists(input_path):
            return {"success": False, "error": f"Input file not found: {input_path}"}

        try:
            ext = os.path.splitext(input_path)[1].lower()
            if ext == ".pdf":
                result = pdf.add_watermark(input_path, watermark_text, output_path, **kwargs)
            elif ext in [".docx", ".doc"]:
                result = docx.add_watermark(input_path, watermark_text, output_path, **kwargs)
            else:
                return {"success": False, "error": f"Watermark not supported for extension: {ext}"}

            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": f"Add watermark failed: {e}"}

    def batch_add_watermark(
        self,
        source_dir: str,
        target_dir: str,
        watermark_text: str,
        **kwargs,
    ) -> Dict[str, Dict[str, Any]]:
        """Batch add watermarks to PDF and DOCX files."""
        if not os.path.exists(source_dir):
            return {"error": f"Source directory not found: {source_dir}"}

        os.makedirs(target_dir, exist_ok=True)
        results: Dict[str, Dict[str, Any]] = {}

        for filename in os.listdir(source_dir):
            ext = os.path.splitext(filename)[1].lower()
            if ext in [".pdf", ".docx"]:
                input_path = os.path.join(source_dir, filename)
                output_path = os.path.join(target_dir, filename)
                results[filename] = self.add_watermark(input_path, watermark_text, output_path, **kwargs)

        return results

    def execute_workflow(self, workflow_config: Dict[str, Any]) -> Dict[str, Any]:
        """Execute a minimal, explicit workflow definition.

        Expected format:
        {
          "steps": [
            {"action": "create", "params": {...}},
            {"action": "convert", "params": {...}},
            ...
          ]
        }
        """
        steps = workflow_config.get("steps", []) if isinstance(workflow_config, dict) else []
        if not isinstance(steps, list):
            return {"success": False, "error": "workflow_config.steps must be a list"}

        step_results = []
        overall_success = True

        for idx, step in enumerate(steps, start=1):
            if not isinstance(step, dict):
                step_results.append({"step": idx, "success": False, "error": "Step must be an object"})
                overall_success = False
                continue

            action = str(step.get("action", "")).strip().lower()
            params = step.get("params", {})
            if not isinstance(params, dict):
                params = {}

            if action == "create":
                result = self.create(**params)
            elif action == "presentation_preflight":
                result = {"success": True, "preferences": pptx.presentation_preflight_wizard()}
            elif action == "convert":
                result = self.convert(**params)
            elif action in {"add_watermark", "watermark"}:
                result = self.add_watermark(**params)
            elif action == "extract_data":
                result = self.extract_data(**params)
            else:
                result = {"success": False, "error": f"Unsupported workflow action: {action}"}

            if not result.get("success", False):
                overall_success = False

            step_results.append({"step": idx, "action": action, **result})

        return {"success": overall_success, "steps": step_results}

    def extract_data(self, input_path: str, **kwargs) -> Dict[str, Any]:
        """Extract text/tabular data from supported document formats."""
        if not os.path.exists(input_path):
            return {"success": False, "error": f"Input file not found: {input_path}"}

        try:
            ext = os.path.splitext(input_path)[1].lower()
            if ext in [".xlsx", ".xls"]:
                data = xlsx.extract_data(input_path, **kwargs)
            elif ext in [".docx", ".doc"]:
                data = docx.extract_text(input_path, **kwargs)
            elif ext == ".pdf":
                data = pdf.extract_text(input_path, **kwargs)
            elif ext == ".pptx":
                data = pptx.extract_text(input_path, **kwargs)
            else:
                return {"success": False, "error": f"Data extraction not supported for extension: {ext}"}

            return {"success": True, "data": data}
        except Exception as e:
            return {"success": False, "error": f"Data extraction failed: {e}"}

    @staticmethod
    def get_presentation_preflight_prompts() -> Dict[str, Any]:
        """Return available preflight options for deck generation."""
        catalog = template_registry.list_templates()
        return {
            "themes": pptx.list_themes(),
            "chart_modes": pptx.list_chart_modes(),
            "tones": pptx.list_tones(),
            "html_themes": html_slides.list_reference_themes(),
            "slide_templates": [t["id"] for t in catalog["slides"]],
            "document_templates": [t["id"] for t in catalog["documents"]],
            "preferences_required_for_strict_mode": ["theme", "chart_mode", "use_emojis", "tone"],
        }

    @staticmethod
    def list_templates(category: Optional[str] = None) -> Dict[str, Any]:
        """List file-based slide/document templates with their slot contracts.

        Args:
            category: optional "slides" or "documents". If omitted, both are returned.
        """
        return template_registry.list_templates(category=category)

    @staticmethod
    def get_template_meta(template_id: str, category: str = "slides") -> Dict[str, Any]:
        """Return the slot contract (meta.json) for a single template."""
        return template_registry.get_template_meta(template_id, category=category)

    @staticmethod
    def _create_single_slide_from_template(**kwargs) -> Dict[str, Any]:
        """Adapter: `doc_type="slide_template"`.

        Expected kwargs:
            template: template id (required)
            data: dict of slot values
            output_path: target path (.png, .pdf or .html)
            output_format: "png" | "pdf" | "html" (default inferred from extension)
        """
        template_id = kwargs.get("template") or kwargs.get("template_id")
        if not template_id:
            raise ValueError("slide_template requires 'template' kwarg")
        output_path = kwargs["output_path"]
        fmt = kwargs.get("output_format") or os.path.splitext(output_path)[1].lstrip(".").lower() or "png"
        data = kwargs.get("data", {}) or {}
        return template_registry.render_slide_template(
            template_id=template_id, data=data,
            output_path=output_path, output_format=fmt,
            width=int(kwargs.get("width", template_registry.SLIDE_WIDTH_PX)),
            height=int(kwargs.get("height", template_registry.SLIDE_HEIGHT_PX)),
        )

    @staticmethod
    def _create_pptx_from_templates(**kwargs) -> Dict[str, Any]:
        """Adapter: `doc_type="slide_deck_templates"`.

        Expected kwargs:
            slides: list of {"template": id, "data": {...}}
            output_path: target .pptx
            keep_renders: bool (optional)
            renders_dir: str (optional)
        """
        slides = kwargs.get("slides")
        if not slides:
            raise ValueError("slide_deck_templates requires non-empty 'slides' list")
        return template_registry.build_pptx_from_templates(
            slides=slides,
            output_path=kwargs["output_path"],
            keep_renders=bool(kwargs.get("keep_renders", False)),
            renders_dir=kwargs.get("renders_dir"),
        )

    @staticmethod
    def _create_document_from_template(**kwargs) -> Dict[str, Any]:
        """Adapter: `doc_type="doc_template"`.

        Expected kwargs:
            template: template id (required)
            data: dict of slot values
            output_path: target path (.pdf, .html or .png)
            output_format: optional override
        """
        template_id = kwargs.get("template") or kwargs.get("template_id")
        if not template_id:
            raise ValueError("doc_template requires 'template' kwarg")
        output_path = kwargs["output_path"]
        fmt = kwargs.get("output_format") or os.path.splitext(output_path)[1].lstrip(".").lower() or "pdf"
        data = kwargs.get("data", {}) or {}
        return template_registry.render_document_template(
            template_id=template_id, data=data,
            output_path=output_path, output_format=fmt,
        )
