#!/usr/bin/env python3
import requests
import sys
import time
import os

API_KEY = "MS183c7da3e93541939f87d04b110e3008"
INPUT_IMAGE = sys.argv[1] if len(sys.argv) > 1 else "input.jpg"
OUTPUT_IMAGE = sys.argv[2] if len(sys.argv) > 2 else "output.jpg"

UPSCALE_URL = "https://api.magnific.ai/v1/upscale"


def upscale_image(input_path: str, output_path: str):
    if not os.path.exists(input_path):
        print(f"Error: file not found: {input_path}")
        sys.exit(1)

    print(f"Uploading {input_path}...")
    with open(input_path, "rb") as f:
        response = requests.post(
            UPSCALE_URL,
            headers={"x-api-key": API_KEY},
            files={"image": (os.path.basename(input_path), f, "image/jpeg")},
            data={
                "type": "Photo",
                "scale": 2,
                "optimizeLowResImages": False,
                "realism": 50,
                "resemblance": 50,
                "fractality": 50,
                "hdr": 50,
            },
        )

    if response.status_code == 202:
        job = response.json()
        job_id = job.get("id")
        print(f"Job submitted: {job_id}. Polling for result...")

        status_url = f"https://api.magnific.ai/v1/upscale/{job_id}"
        while True:
            time.sleep(5)
            r = requests.get(status_url, headers={"x-api-key": API_KEY})
            data = r.json()
            status = data.get("status")
            print(f"  Status: {status}")
            if status == "done":
                image_url = data.get("resultUrl") or data.get("url")
                img_data = requests.get(image_url).content
                with open(output_path, "wb") as out:
                    out.write(img_data)
                print(f"Saved upscaled image to: {output_path}")
                break
            elif status in ("failed", "error"):
                print(f"Job failed: {data}")
                sys.exit(1)

    elif response.status_code == 200:
        with open(output_path, "wb") as out:
            out.write(response.content)
        print(f"Saved upscaled image to: {output_path}")

    else:
        print(f"Error {response.status_code}: {response.text}")
        sys.exit(1)


if __name__ == "__main__":
    upscale_image(INPUT_IMAGE, OUTPUT_IMAGE)
