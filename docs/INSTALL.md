# Installation | Instalación

## English

### Requirements
- Python 3.8+
- Linux / macOS / Windows
- Optional: LibreOffice for high-quality format conversion

### 1) Clone and install

```bash
git clone https://github.com/rafalozan0/DocFlow-Presentations-and-Docs-Skill.git
cd DocFlow-Presentations-and-Docs-Skill
pip install -r requirements.txt
```

### 2) Optional system tools

Linux (Ubuntu/Debian):
```bash
sudo apt update
sudo apt install -y libreoffice
```

macOS:
```bash
brew install --cask libreoffice
```

Windows:
- Install LibreOffice from https://www.libreoffice.org/

### 3) Verify

```bash
python -m compileall src examples
python examples/basic_usage.py
python examples/presentation_style_chart_demo.py
```

### 4) Isolated runtime option (uv)

If your environment has no pip or you want ephemeral dependencies:

```bash
uv run --with python-pptx --with python-docx --with openpyxl --with reportlab --with pypdf2 --with pandas --with matplotlib python examples/basic_usage.py
```

## Español

### Requisitos
- Python 3.8+
- Linux / macOS / Windows
- Opcional: LibreOffice para conversiones de alta calidad

### 1) Clonar e instalar

```bash
git clone https://github.com/rafalozan0/DocFlow-Presentations-and-Docs-Skill.git
cd DocFlow-Presentations-and-Docs-Skill
pip install -r requirements.txt
```

### 2) Herramientas del sistema (opcional)

Linux (Ubuntu/Debian):
```bash
sudo apt update
sudo apt install -y libreoffice
```

macOS:
```bash
brew install --cask libreoffice
```

Windows:
- Instalar LibreOffice desde https://www.libreoffice.org/

### 3) Verificación

```bash
python -m compileall src examples
python examples/basic_usage.py
python examples/presentation_style_chart_demo.py
```

### 4) Opción de runtime aislado (uv)

Si tu entorno no tiene pip o prefieres dependencias efímeras:

```bash
uv run --with python-pptx --with python-docx --with openpyxl --with reportlab --with pypdf2 --with pandas --with matplotlib python examples/basic_usage.py
```
