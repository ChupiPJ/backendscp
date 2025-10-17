from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import io

from .models import RenderRequest
from .ppt import generate_presentation

TEMPLATE_PATH = str(Path(__file__).parent / "templates" / "silicon_eic_template.pptx")

app = FastAPI()

# Ajusta orígenes si quieres restringirlo
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _build_replacements(req: RenderRequest) -> dict:
    """
    Solo usamos los placeholders que tu plantilla necesita.
    Enviamos las llaves con {{...}} para que el replace sea exacto.
    """
    repl = {
        "{{COMPANY_NAME}}": req.company_name,
    }

    # pricing_overrides: solo añadimos los que vengan
    po = req.pricing_overrides or {}
    if po.SETUP_FEE is not None:
        repl["{{SETUP_FEE}}"] = str(po.SETUP_FEE)
    if po.SHORT_FEE is not None:
        repl["{{SHORT_FEE}}"] = str(po.SHORT_FEE)
    if po.FULL_FEE is not None:
        repl["{{FULL_FEE}}"] = str(po.FULL_FEE)
    if po.GRANT_FEE is not None:
        repl["{{GRANT_FEE}}"] = po.GRANT_FEE
    if po.EQUITY_FEE is not None:
        repl["{{EQUITY_FEE}}"] = po.EQUITY_FEE

    return repl


@app.post("/render")
async def render_presentation(request: RenderRequest):
    try:
        replacements = _build_replacements(request)
        buf = generate_presentation(
            template_path=TEMPLATE_PATH,
            replacements=replacements,
            slide_toggles=request.slide_toggles or {},
        )
        filename = f"proposal_{request.company_name.replace(' ', '_')}.pptx"
        return StreamingResponse(
            io.BytesIO(buf.getvalue()),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
