from flask import (Flask, render_template, request,
        redirect, send_from_directory, send_file)
import os
import pandas as pd
from glob import glob
from io import BytesIO
from zipfile import ZipFile
from pydub import AudioSegment
from pydub.silence import split_on_silence

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    chunked_files = []
    folder_name = ''
    if request.method == "POST":
        if "audioFile" not in request.files:
            return redirect(request.url)

        audioFile = request.files["audioFile"]
        nameFile = request.files["nameFile"]
        audioFileName = audioFile.filename
        if audioFile.filename == "":
            return redirect(request.url)

        if audioFile:
            df = pd.read_excel(nameFile, index_col=None, na_values=['NA'],
                                usecols="A", header=None)
            df = df.dropna()
            file_names = []
            for column in df.columns:
                li = df[column].tolist()
                file_names.append(li)

            sound = AudioSegment.from_file(audioFile)
            audio_chunks = split_on_silence(sound, min_silence_len=200, silence_thresh=-60)
            folder_name = audioFileName.split('.')[0]
            file_format = audioFileName.split('.')[1]

            if not os.path.isdir(folder_name):
                os.mkdir(folder_name)

            for i, chunk in enumerate(audio_chunks):
                try:
                    fileName = f'{file_names[0][i]}.{file_format}'
                    output_file = os.path.join(folder_name, fileName)
                    chunk.export(output_file, format=file_format)
                    chunked_files.append(fileName)
                except IndexError:
                    pass

    return render_template('index.html', folder_name=folder_name, chunked_files=chunked_files)


@app.route("/download/<path:folder_name>/<path:filename>")
def download(folder_name, filename):
    return send_from_directory(f'{folder_name}/', filename, as_attachment=True)


@app.route("/download_all/<path:folder_name>")
def download_all(folder_name):
    stream = BytesIO()
    with ZipFile(stream, 'w') as zf:
        for file in glob(os.path.join(folder_name, '*')):
            zf.write(file, os.path.basename(file))
    stream.seek(0)

    return send_file(stream, as_attachment=True, download_name=f'{folder_name}.zip')


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
