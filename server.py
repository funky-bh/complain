"""Flask web server for the ChannelTalk CS complaint analyzer."""

import os
import tempfile

from dotenv import load_dotenv
from flask import Flask, render_template, request, send_file

load_dotenv()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze_file():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return render_template("index.html", error="파일을 선택해주세요.")

    if not uploaded.filename.lower().endswith(".xlsx"):
        return render_template("index.html", error=".xlsx 파일만 업로드할 수 있습니다.")

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        return render_template("index.html", error="ANTHROPIC_API_KEY가 설정되지 않았습니다. .env 파일을 확인해주세요.")

    tmp_in = tmp_out = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            uploaded.save(f)
            tmp_in = f.name

        tmp_out = tmp_in.replace(".xlsx", "_analyzed.xlsx")

        from analyze import analyze
        analyze(tmp_in, tmp_out)

        original_stem = uploaded.filename.rsplit(".", 1)[0]
        return send_file(
            tmp_out,
            as_attachment=True,
            download_name=f"{original_stem}_analyzed.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except ValueError as e:
        return render_template("index.html", error=str(e))
    except Exception as e:
        return render_template("index.html", error=f"분석 중 오류가 발생했습니다: {e}")
    finally:
        for path in (tmp_in, tmp_out):
            if path:
                try:
                    os.unlink(path)
                except FileNotFoundError:
                    pass


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)
