from flask import Flask, render_template, request, jsonify
import os
import subprocess
import uuid

app = Flask(__name__)
UPLOAD_FOLDER = "temp_files"
ALLOWED_EXTENSIONS = {"pptx"}

# Ensure the temp folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    if file and allowed_file(file.filename):
        unique_filename = f"{uuid.uuid4()}.pptx"
        file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
        file.save(file_path)

        return jsonify({"filename": file.filename, "filepath": file_path}), 200

    return jsonify({"error": "Invalid file type"}), 400

# Global variable to store the process reference for the running presentation
presentation_process = None

@app.route("/start", methods=["POST"])
def start_presentation():
    global presentation_process

    data = request.get_json()
    file_path = data.get("filepath")
    if not file_path or not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 400

    # Run Slide.py with the uploaded file
    try:
        presentation_process = subprocess.Popen(["python", "Slide.py", file_path])
        return jsonify({"message": "Presentation started"}), 200
    except Exception as e:
        return jsonify({"error": f"Failed to start presentation: {str(e)}"}), 500

@app.route("/stop", methods=["POST"])
def stop_presentation():
    global presentation_process
    if presentation_process:
        # Terminate the presentation process
        presentation_process.terminate()
        presentation_process = None

        # Clean up the uploaded file
        for file in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, file)
            os.remove(file_path)

        return jsonify({"message": "Presentation stopped, file deleted"}), 200

    return jsonify({"error": "No active presentation to stop"}), 400

if __name__ == "__main__":
    app.run(debug=True)
