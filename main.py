# main.py (REPLACE your existing file with this)
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

app = FastAPI(title="Excel + Social Agent MVP (Improved Excel Reader)")

def infer_header_and_read_excel(bytes_data):
    """
    Read excel robustly: try normal read; if header issues, read with header=None,
    find first row with >1 non-null and use it as header.
    """
    try:
        # try normal read first
        df = pd.read_excel(io.BytesIO(bytes_data), engine="openpyxl")
        # if dataframe has mostly NaNs in header or first row empty, fallback below
        if df.shape[0] == 0 or df.shape[1] == 0:
            raise Exception("empty")
        return df
    except Exception:
        # fallback: read without header, find header row
        tmp = pd.read_excel(io.BytesIO(bytes_data), header=None, engine="openpyxl")
        # drop fully empty rows
        tmp = tmp.dropna(how="all").reset_index(drop=True)
        if tmp.shape[0] == 0:
            raise
        # find first row which has more than one non-null value -> candidate header
        header_row_idx = None
        for i in range(0, min(5, len(tmp))):
            non_null_count = tmp.iloc[i].count()
            if non_null_count >= 1:
                header_row_idx = i
                break
        if header_row_idx is None:
            header_row_idx = 0
        # set header
        new_header = tmp.iloc[header_row_idx].astype(str).str.strip().tolist()
        data = tmp.iloc[header_row_idx+1 : ].reset_index(drop=True)
        data.columns = new_header
        return data

def normalize_column_name(col):
    return str(col).strip()

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # Drop fully empty columns and rows
    df = df.dropna(axis=0, how='all').dropna(axis=1, how='all').reset_index(drop=True)

    # Normalize column names
    df.columns = [normalize_column_name(c) for c in df.columns]

    # strip strings
    for c in df.select_dtypes(include=["object"]).columns:
        df[c] = df[c].astype(str).str.strip()

    # Try to parse date-like columns
    for col in df.columns:
        if col and col.lower() in ("date", "transaction date", "txn_date", "datetime"):
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d-%m-%Y")
            except Exception:
                pass

    # Try to find amount-like column and convert to numeric
    cols_lower = [c.lower() for c in df.columns]
    amount_col = None
    for candidate in ("amount", "amt", "price", "total", "value"):
        if candidate in cols_lower:
            amount_col = df.columns[cols_lower.index(candidate)]
            break
    if amount_col:
        # remove currency symbols, commas
        df[amount_col] = df[amount_col].astype(str).str.replace(r"[^\d\.\-]", "", regex=True)
        df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)

    # Drop duplicate rows (exact duplicates)
    df = df.drop_duplicates().reset_index(drop=True)

    # Reset index and return
    return df

@app.post("/upload-and-clean")
async def upload_and_clean(file: UploadFile = File(...)):
    contents = await file.read()
    filename = file.filename or "cleaned.xlsx"

    # read file robustly (xlsx or csv)
    df = None
    try:
        # try reading as excel robustly
        df = infer_header_and_read_excel(contents)
    except Exception:
        try:
            # try csv
            df = pd.read_csv(io.StringIO(contents.decode("utf-8")))
        except Exception as e:
            return JSONResponse({"error": "File read error. Upload .xlsx/.xls/.csv. Details: " + str(e)}, status_code=400)

    cleaned = clean_dataframe(df)

    # create top5 if possible
    top5 = []
    cols_lower = [c.lower() for c in cleaned.columns]
    if "product" in cols_lower and any(x in cols_lower for x in ("amount", "amt", "price", "total", "value")):
        prod_col = cleaned.columns[cols_lower.index("product")]
        amt_idx = next(i for i, x in enumerate(cols_lower) if x in ("amount", "amt", "price", "total", "value"))
        amt_col = cleaned.columns[amt_idx]
        grouped = cleaned.groupby(prod_col)[amt_col].sum().reset_index().sort_values(by=amt_col, ascending=False).head(5)
        top5 = grouped.to_dict(orient="records")

    # write cleaned to bytes with same extension
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        cleaned.to_excel(writer, index=False, sheet_name="cleaned")
    out.seek(0)

    headers = {"Content-Disposition": f"attachment; filename=cleaned_{filename}"}
    # Also return a small JSON in header not possible; keep download response
    return StreamingResponse(out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers=headers)

@app.post("/clean-and-summary")
async def clean_and_summary(file: UploadFile = File(...)):
    contents = await file.read()
    try:
        df = infer_header_and_read_excel(contents)
    except Exception:
        try:
            df = pd.read_csv(io.StringIO(contents.decode("utf-8")))
        except Exception as e:
            return JSONResponse({"error": "File read error. Upload .xlsx/.xls/.csv. Details: " + str(e)}, status_code=400)
    cleaned = clean_dataframe(df)
    # Provide useful summary
    cols_lower = [c.lower() for c in cleaned.columns]
    top5 = []
    if "product" in cols_lower and any(x in cols_lower for x in ("amount", "amt", "price", "total", "value")):
        prod_col = cleaned.columns[cols_lower.index("product")]
        amt_idx = next(i for i, x in enumerate(cols_lower) if x in ("amount", "amt", "price", "total", "value"))
        amt_col = cleaned.columns[amt_idx]
        grouped = cleaned.groupby(prod_col)[amt_col].sum().reset_index().sort_values(by=amt_col, ascending=False).head(5)
        top5 = grouped.to_dict(orient="records")

    return {
        "rows": len(cleaned),
        "columns": list(cleaned.columns),
        "top5_products_by_amount": top5
    }

@app.post("/generate-caption")
async def generate_caption(product_name: str = Form(...), price: str = Form(""), feature: str = Form(""), tone: str = Form("friendly")):
    if not OPENAI_KEY:
        short = f"{product_name} ab available! Starting at {price}. {feature}"
        long = f"Introducing {product_name} — {feature}. Grab yours at just {price}. Limited stock!"
        return {"short_caption": short, "long_caption": long, "note": "OpenAI key not set; returned simple templates."}

    prompt = f"""
    You are a helpful social media copywriter writing in Hindi (and a little English if helpful).
    Product: {product_name}
    Price: {price}
    Feature: {feature}
    Tone: {tone}

    Provide:
    1) Short caption (max 70 characters)
    2) Long caption (100-220 characters)
    3) 8 relevant hashtags (comma separated)
    Return the response in JSON format with keys: short, long, hashtags
    """

    try:
        completion = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=300,
            temperature=0.7
        )
        text = completion.choices[0].message.content.strip()
        import json
        try:
            parsed = json.loads(text)
            return parsed
        except Exception:
            return {"raw": text}
    except Exception as e:
        return JSONResponse({"error": "OpenAI API error", "details": str(e)}, status_code=500)

@app.get("/", response_class=HTMLResponse)
def homepage():
    return """
    <h2>Excel + Social Agent (MVP)</h2>
    <p>Use /upload-and-clean to upload an Excel file (returns cleaned excel)</p>
    <p>Use /clean-and-summary to get JSON summary (top5 products etc.)</p>
    <p>Use /generate-caption to create captions (requires OPENAI_API_KEY)</p>
    <hr/>
    <p>Example HTML Test form available at <a href="/test">/test</a></p>
    """

@app.get("/test", response_class=HTMLResponse)
def test_page():
    html = """
    <h3>Upload Excel and get cleaned file</h3>
    <form action="/upload-and-clean" enctype="multipart/form-data" method="post">
      <input name="file" type="file" />
      <input type="submit" value="Upload and Clean" />
    </form>
    <hr/>
    <h3>Generate Caption (Form)</h3>
    <form action="/generate-caption" method="post">
      Product Name: <input name="product_name" /><br/>
      Price: <input name="price" /><br/>
      Feature: <input name="feature" /><br/>
      Tone: <select name="tone"><option>friendly</option><option>formal</option><option>humorous</option></select><br/>
      <input type="submit" value="Generate Caption" />
    </form>
    """
    return HTMLResponse(content=html)
