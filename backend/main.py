from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from compare import compare_excels, compare_newflow_bytes

app = FastAPI(title="SR Compare API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-New-SR-Count", "Content-Disposition"],
)

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/compare")
async def compare(
    old_file: UploadFile = File(...),
    new_file: UploadFile = File(...),
    sheet_name: str = Form(None),
):
    try:
        out_bytes, count = compare_excels(
            await old_file.read(),
            await new_file.read(),
            sheet_name,
        )

        headers = {
            "Content-Disposition": 'attachment; filename="YENI_GELEN_SR.xlsx"',
            "X-New-SR-Count": str(count),
        }
        return Response(
            content=out_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})

@app.post("/compare-newflow")
async def compare_newflow(
    old_file: UploadFile = File(...),
    new_file: UploadFile = File(...),
    sheet_name: str = Form(None),
):
    try:
        out_bytes, count = compare_newflow_bytes(
            await old_file.read(),
            await new_file.read(),
            sheet_name,
        )

        headers = {
            "Content-Disposition": 'attachment; filename="NEW_FLOW_SECUNET_TPBE.xlsx"',
            "X-New-SR-Count": str(count),
        }
        return Response(
            content=out_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})
