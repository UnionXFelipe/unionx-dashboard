"""
Hook: se ejecuta en cada evento Stop de Claude Code.
1. Si hay una seccion "## Pendiente" en CLAUDE.md, la integra al contenido principal y la elimina.
2. Actualiza siempre la linea "Ultima actualizacion" con la fecha/hora actual.
"""
import sys, re
from datetime import datetime
from pathlib import Path

CLAUDEMD = Path(r'C:\Users\felip\Desktop\UnionX Cloude\CLAUDE.md')

if not CLAUDEMD.exists():
    sys.exit(0)

content = CLAUDEMD.read_text(encoding='utf-8')
now = datetime.now().strftime('%Y-%m-%d %H:%M')

# ── 1. Actualizar o insertar linea de ultima actualizacion ─────────────────
UPD_PATTERN = re.compile(r'^> \*\*Ultima actualizacion:\*\*.*$', re.MULTILINE)
UPD_LINE    = f'> **Ultima actualizacion:** {now}'

if UPD_PATTERN.search(content):
    content = UPD_PATTERN.sub(UPD_LINE, content)
else:
    # Insertar justo despues del titulo principal (primera linea #)
    content = re.sub(r'(^# .+$)', r'\1\n\n' + UPD_LINE, content, count=1, flags=re.MULTILINE)

# ── 2. Integrar seccion "## Pendiente" si existe ───────────────────────────
PENDING_PATTERN = re.compile(
    r'\n## Pendiente\n(.*?)(?=\n## |\Z)',
    re.DOTALL
)
m = PENDING_PATTERN.search(content)
if m:
    pending_body = m.group(1).strip()
    # Eliminar la seccion pendiente del contenido
    content = PENDING_PATTERN.sub('', content)
    # Agregar el contenido pendiente al final de la seccion de errores conocidos
    # o al final del archivo si no existe esa seccion
    ERRORS_HEADER = '## Errores conocidos y sus soluciones'
    if ERRORS_HEADER in content:
        # Insertar antes del siguiente ## o al final
        content = content.rstrip() + f'\n\n### Actualizacion {now}\n{pending_body}\n'
    else:
        content = content.rstrip() + f'\n\n## Actualizaciones recientes\n\n### {now}\n{pending_body}\n'

CLAUDEMD.write_text(content, encoding='utf-8')
