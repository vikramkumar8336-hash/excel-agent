import io
import os
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse
import pandas as pd
from dotenv import load_dotenv
import openai

load_dotenv()
OPENAI_KEY = os.getenv("OPENAI_API_KEY")
if OPENAI_KEY:
    openai.api_key = OPENAI_KEY

app = FastAPI(title="Excel + Social Agent MVP")

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(col).strip() for col in df.columns]
    df = df.drop_duplicates().reset_index(drop=True)
    for c in df.select_dtypes(include=["object"]).columns:
        df[c] = df[c].astype(str).str.strip()
    for col in df.columns:
        if col.lower() in ("date", "transaction date", "txn_date", "datetime"):
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d-%m-%Y")
            except:
                pass
    for possible in ("amount", "amt", "price", "total", "value"):
        if possible in [c.lower() for c in df.columns]:
            for c in df.columns:
                if c.lower() == possible:
                    df[c] = pd.to_numeric(
                        df[c].astype(str).str.replace(",", "").str.replace("₹", "").str.strip(),
                        errors="coerce"
                    ).fillna(0)
    return df

@app.post("/upload-and-clean")
async def upload_and_clean(file: UploadFile = File(...)):
    contents = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(contents))
    except:
        try:
            df = pd.read_csv(io.StringIO(contents.decode("utf-8")))
        except:
            return JSONResponse({"error": "Invalid file"}, status_code=400)

    cleaned = clean_dataframe(df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned.to_excel(writer, index=False, sheet_name="cleaned")
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=cleaned.xlsx"}
    )

@app.post("/clean-and-summary")
async def clean_and_summary(file: UploadFile = File(...)):
    contents = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(contents))
    except:
        try:
            df = pd.read_csv(io.StringIO(contents.decode("utf-8")))
        except:
            return JSONResponse({"error": "Invalid file"}, status_code=400)

    cleaned = clean_dataframe(df)

    return {
        "rows": len(cleaned),
        "columns": list(cleaned.columns)
    }

@app.post("/generate-caption")
async def generate_caption(
    product_name: str = Form(...),
    price: str = Form(""),
    feature: str = Form("")
):
    if not OPENAI_KEY:
        return {
            "short": f"{product_name} now available!",
            "long": f"{product_name} — {feature}. Price: {price}",
            "note": "OpenAI key missing, simple captions generated."
        }

    prompt = f"""
    Write Instagram captions in Hinglish.
    Product: {product_name}
    Price: {price}
    Feature: {feature}
    
    Return:
    short_caption:
    long_caption:
    hashtags:
    """

    res = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}]
    )

    return {"result": res.choices[0].message.content}

@app.get("/", response_class=HTMLResponse)
def home():
    return "<h2>Excel + Social AI Agent Running!</h2><a href='/test'>Go to Test Page</a>"

@app.get("/test", response_class=HTMLResponse)
def test_page():
    return """
    <h3>Upload Excel for Cleaning</h3>
    <form method="post" action="/upload-and-clean" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">Upload & Clean</button>
    </form>

    <hr>

    <h3>Generate Social Caption</h3>
    <form method="post" action="/generate-caption">
        <input name="product_name" placeholder="Product Name"><br>
        <input name="price" placeholder="Price"><br>
        <input name="feature" placeholder="Feature"><br>
        <button type="submit">Generate Caption</button>
    </form>
    """
