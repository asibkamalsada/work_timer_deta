from fastapi import FastAPI, UploadFile, Request, Form
from fastapi.templating import Jinja2Templates
from starlette.responses import Response

import work_timer

app = FastAPI()
templates = Jinja2Templates(directory="templates/")


@app.get("/")
async def root(request: Request):
    return templates.TemplateResponse("root.html", {"request": request})


@app.post("/uploadfile/")
async def create_upload_file(file: UploadFile = Form()):
    name, binary = await work_timer.convert(file)
    return Response(content=binary, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})
