from flask import Flask, send_file
from flask_cors import CORS, cross_origin
from pptx import Presentation
import collections.abc

app = Flask(__name__)
cors = CORS(app)

# Create a presentation object
presentation = Presentation()

# Add a slide with a title and content
slide_layout = presentation.slide_layouts[
  1]  # Choose a slide layout (e.g., Title and Content)
slide = presentation.slides.add_slide(slide_layout)

title = slide.shapes.title
title.text = "My First Slide"

content = slide.placeholders[1]
content.text = "Hello, PowerPoint!"

# Save the presentation to a file
presentation.save("my_presentation.pptx")


@app.route('/')
@cross_origin()
def index():
  file_path = 'my_presentation.pptx'

  return send_file(file_path, as_attachment=True)


app.run(host='0.0.0.0', port=5000, debug=True)
