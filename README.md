# Invoice → Fakturama

## Requirements

- **Windows.** The app-driving side uses UIA via `pywinauto` and `pywin32`.
- **Python 3.10–3.12.** `paddlepaddle` 2.6.x publishes no wheels for 3.13+.
- **Fakturama**, running, with the **English** UI. Controls are located by
  visible name, so a differently-localised install needs its own name map.

## Setup

```bash
uv venv --python 3.11 .venv
.venv\Scripts\activate
uv pip install -r requirements.txt
```

`requirements.txt` pins `paddleocr<3` and `paddlepaddle==2.6.2` deliberately —
newer releases miscompile the bundled OCR models. See the comments in that file
before bumping anything.

## Usage

```bash
python main.py                  # read invoice.pdf
python main.py path/to/doc.pdf  # read another document
python main.py --cursor         # draw an on-screen pointer so you can watch
```

Watching a run can also be enabled with `FAKTURAMA_TRACE=visual`. It is off by
default so unattended runs stay fast. The overlay is drawn — it never moves the
real mouse or keyboard.

```bash
python smoke.py     # the same steps on hardcoded demo data, no OCR (~1 min)
```

`smoke.py` is the fast check of the automation half on its own: `main.py`
spends ~25s in OCR before it touches the app, and `smoke.py` skips that by
using the values extraction returns for `invoice.pdf`.

## Modules

| Module | Responsibility                                                                                                        |
|---|-----------------------------------------------------------------------------------------------------------------------|
| `main.py` | Orchestrates the five stages; maps extracted fields to UI fields                                                      |
| `smoke.py` | The same run on demo values, without OCR                                                                              |
| `fakturama.py` | Everything that talks to the application                                                                              |
| `table.py` | Reads Fakturama's drawn lists (they have no accessibility tree) by OCR-ing a capture and cutting it on the grid lines |
| `tracing.py` | The optional on-screen pointer and the step log                                                                       |
| `sales_order.py` | The document layout: which label means which field                                                                    |
| `layout.py` | Geometry, grouping positioned text into rows and columns                                                              |
| `ocr.py` | The only module that knows about PaddleOCR                                                                            |
| `models.py` | The `SalesOrder` schema, and the rules over it (address parsing, gross price, line totals)                            |
