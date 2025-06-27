# PowerPoint Generator

An intelligent PowerPoint presentation generator that automatically creates and modifies PowerPoint presentations based on templates and outlines. Built with Python and OpenAI's API.

## Features

- **Template-Based Generation**: Uses predefined PowerPoint templates to maintain consistent styling
- **Smart Content Generation**: Leverages AI to generate appropriate content for different slide types
- **Flexible Outline Structure**: Supports customizable presentation outlines in JSON format
- **Automatic Text Placement**: Maps content to specific text boxes using template descriptions
- **Chart and SmartArt Support**: Handles complex PowerPoint elements including charts and SmartArt
- **Multilingual Support**: Capable of generating content in multiple languages

## Prerequisites

- Python 3.13+
- OpenAI API key
- Required Python packages (see `requirements.txt`)

## Installation

1. Clone the repository:

```bash
git clone https://github.com/CY202227/PPT_GENERATE.git
cd PPT_Generate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root and add your API credentials:

```
API_KEY=your_openai_api_key
BASE_URL=your_api_base_url
```

## Usage

1. Prepare your template files:

   - Place your PowerPoint template in the project root as `template.pptx`
   - Create a template description file `template_description.json`
   - Create an outline file `template_outline_01.json`
2. Run the generator:

```bash
python PowerPoint_generator.py
```

## Template Description Format

The `template_description.json` file defines the structure of your PowerPoint template:

```json
{
  "slide01": {
    "type": "title_slide",
    "description": "Title Slide",
    "text_elements": {
      "main_title": {
        "id": 3,
        "description": "Main title text box"
      }
      // ... more elements
    }
  }
  // ... more slides
}
```

## Outline Format

The `template_outline_01.json` file defines your presentation content:

```json
{
  "main_title": "Your Presentation Title",
  "speaker_name": "Presenter Name",
  "chapter01_title": "Chapter 1",
  "sections": [
    {
      "section01": {
        "title": "Section 1 Content"
      }
      // ... more sections
    }
  ]
  // ... more chapters
}
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- OpenAI for providing the API
- Python-PPTX library for PowerPoint manipulation
- All contributors and users of this project
