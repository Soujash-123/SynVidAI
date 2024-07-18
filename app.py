from flask import Flask, request, render_template, send_file, after_this_request
import os
import tempfile
from docx import Document
from PIL import Image
from gtts import gTTS
from moviepy.editor import ImageSequenceClip, AudioFileClip, concatenate_videoclips

app = Flask(__name__)


def extract_images_and_text(doc_path, temp_dir):
    doc = Document(doc_path)
    images = []
    text = []

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image_data = rel.target_part.blob
            image_name = os.path.join(temp_dir, f'image{len(images) + 1}.png')
            with open(image_name, 'wb') as f:
                f.write(image_data)
            images.append(image_name)

    for para in doc.paragraphs:
        text.append(para.text)

    return images, text


def create_voice_over(text, audio_path):
    full_text = " ".join(text)
    tts = gTTS(text=full_text, lang='en')
    tts.save(audio_path)


def create_slideshow(images, audio_path, video_path, default_duration=50):
    audio = AudioFileClip(audio_path)
    clip_duration = audio.duration

    if clip_duration < 150:
        duration_per_image = clip_duration / len(images)
    else:
        duration_per_image = default_duration

    clips = []
    for img in images:
        img_clip = ImageSequenceClip([img], durations=[duration_per_image])
        clips.append(img_clip)

    video = concatenate_videoclips(clips, method="compose")
    video = video.set_audio(audio)
    video.write_videofile(video_path, fps=24)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.docx'):
            temp_dir = tempfile.mkdtemp()
            try:
                doc_path = os.path.join(temp_dir, file.filename)
                file.save(doc_path)
                audio_path = os.path.join(temp_dir, 'voiceover.mp3')
                video_path = os.path.join(temp_dir, 'slideshow.mp4')

                images, text = extract_images_and_text(doc_path, temp_dir)
                create_voice_over(text, audio_path)
                create_slideshow(images, audio_path, video_path)

                @after_this_request
                def cleanup(response):
                    try:
                        if os.path.exists(temp_dir):
                            for root, dirs, files in os.walk(temp_dir, topdown=False):
                                for name in files:
                                    os.remove(os.path.join(root, name))
                                for name in dirs:
                                    os.rmdir(os.path.join(root, name))
                            os.rmdir(temp_dir)
                    except Exception as e:
                        print(f"Error cleaning up: {e}")
                    return response

                return send_file(video_path, as_attachment=True)
            except Exception as e:
                print(f"Error processing the file: {e}")
                return "An error occurred during processing."
    return render_template('index.html')


if __name__ == "__main__":
    app.run(debug=True)
