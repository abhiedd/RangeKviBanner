# Banners Multi-Tab Output App (with Images)

A Streamlit app for processing Banners Excel sheets and a product CSV, generating ready-to-use campaign banner outputs (and downloadable images).

## Features

- Upload your Banners Excel (with KVI, Range, Dual MRP tabs).
- Upload a product CSV with `MB_id` and `image_src`.
- Output: Three tabs (KVI, Range, Dual MRP) with:
  - Banner product columns (Product Name, MB ID, Focused Sub Cat, Copy if needed)
  - Img1/Img2: Milkbasket product links (fetched by MB ID)
  - AmzId1/AmzId2: Figma S3 links (.png ext, by MB ID)
- Download multi-tab Excel with all data.
- Download all unique images as PNG zip.
- Download all images with background removed (via rembg).

The necessary colums needs to be present in the excel sheet, in the following format - Hubs | Product Name | Focused Sub Cat |	MB ID 1 | MB ID 2 |	Banner Call-Out (For Range)

## How to run

```bash
pip install -r requirements.txt
streamlit run main.py
