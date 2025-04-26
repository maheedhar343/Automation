

# Document Generator


**Python**
**Flask**
**License**

**A web-based application that generates Word documents from an Excel sheet and a Word template, with support for embedding images from a user-uploaded folder. The project uses Flask for the front-end and back-end, allowing users to upload files through a simple web interface and download the generated document.**


**Prerequisites**

* **Python 3.8+**: Ensure Python is installed on your system.
* **pip**: Python package manager to install dependencies.
* **Google Chrome or Edge**: Recommended browsers for folder uploads (due to **webkitdirectory** support).

  ```
  document-generator/
  │
  ├── static/
  │   ├── css/
  │   │   └── styles.css       # CSS for styling the front-end
  │   └── js/
  │       └── script.js        # JavaScript for front-end interactivity
  │
  ├── templates/
  │   └── index.html           # HTML template for the front-end
  │
  ├── uploads/                 # Temporary folder for the uploaded Excel file
  │   └── data_ples.xlsx       # Single Excel sheet (created after upload)
  │
  ├── generated_docs/          # Folder for the Word template
  │   └── document_1.docx      # Word template (created after upload)
  │
  ├── path/                    # Folder for uploaded images
  │   ├── to/
  │   │   ├── image1.png       # Example image
  │   │   └── image2.jpg       # Example image
  │   └── another_image.jpg    # Example image
  │
  ├── Final_output.docx        # Generated output (created after generation)
  ├── styling.py               # Utility functions for document styling
  ├── generate_document.py     # Backend script for document generation
  ├── app.py                   # Flask app for front-end and back-end integration
  ├── requirements.txt         # List of Python dependencies
  └── README.md                # Project documentation
  ```



**Usage**

1. **Run the Application**:
   Start the Flask development server

```bash
python app.py
```


**These are listed in **requirements.txt**:**


```*
flask
pandas
python-docx
openpyxl
```




**Known Limitations**

* **Browser Support for Folder Uploads**: The **webkitdirectory** attribute for folder uploads is supported in Chrome, Edge, and Opera. Firefox support is limited, and other browsers may not support folder uploads.
* **Image Paths**: The Excel sheet must reference images with paths starting with **path/** (e.g., **path/to/image1.png**). Ensure the uploaded **path/** folder matches this structure.
* **Error Handling**: Basic validation is included, but you may encounter errors if the Excel sheet or template is missing required columns or if image paths are invalid.
