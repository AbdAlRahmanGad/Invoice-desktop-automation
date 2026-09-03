# Invoice → Fakturama

## Requirements

- **Windows.** The app-driving side uses UIA via `pywinauto` and `pywin32`.
- **Python 3.10–3.12.** `paddlepaddle` 2.6.x publishes no wheels for 3.13+.
- **Fakturama**, running, with the **English** UI. Controls are located by
  visible name, so a differently-localised install needs its own name map.

## Setup

```bash
uv venv --python 3.11 .venv && source .venv/bin/activate
uv pip install -r requirements.txt
```

`requirements.txt` pins `paddleocr<3` and `paddlepaddle==2.6.2` deliberately —
newer releases miscompile the bundled OCR models. See the comments in that file
before bumping anything.

## Usage

```bash
python main.py                 # read invoice.pdf
python main.py path/to/doc.pdf  # read another document
python main.py --cursor         # draw an on-screen pointer so you can watch
```

Watching a run can also be enabled with `FAKTURAMA_TRACE=visual`. It is off by
default so unattended runs stay fast. The overlay is drawn — it never moves the
real mouse or keyboard.

Each module runs on its own as a smoke test:

```bash
python fakturama.py   # drive the app with hardcoded demo data (no OCR)
```
