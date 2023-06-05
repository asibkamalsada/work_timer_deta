from fastapi import FastAPI, UploadFile, Request, Form
from fastapi.templating import Jinja2Templates
from starlette.background import BackgroundTask
from starlette.responses import FileResponse, RedirectResponse

import work_timer

app = FastAPI()
templates = Jinja2Templates(directory="templates/")


@app.get("/")
async def root(request: Request):
    return templates.TemplateResponse("root.html", {"request": request})


@app.post("/uploadfile/")
async def create_upload_file(file: UploadFile = Form()):
    if file.size > 0:
        name, file_path = await work_timer.convert(file)
        return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            filename=name, background=BackgroundTask(work_timer.cleanup, file_path))
    else:
        return RedirectResponse("/", status_code=303)
