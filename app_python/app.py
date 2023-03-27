from fastapi import FastAPI, Request, File, UploadFile
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
#pacotes próprios
from manipulacaoPlanilha import gerar_zipfile, receber_planilha, gerar_planilhas

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get('/', response_class=HTMLResponse)
async def main(request:Request):
    return templates.TemplateResponse('home.html', {'request':request})
    

@app.post("/outro")
def image_filter(img: UploadFile = File(...)):
    original_image = Image.open(img.file)
    original_image = original_image.filter(ImageFilter.BLUR)

    filtered_image = BytesIO()
    original_image.save(filtered_image, "JPEG")
    filtered_image.seek(0)

    return StreamingResponse(filtered_image, media_type="image/jpeg")
    

@app.post('/analisar')
async def analisar(request):
    arquivo = request
    
    #processa as planilhas
    gerar_planilhas(receber_planilha())
    zipfile = gerar_zipfile()
    
    return FileResponse(zipfile, status_code = 200)
