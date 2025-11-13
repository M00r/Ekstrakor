# Ekstrakor
Overview
This is a Tkinter-based desktop tool that collects media from chosen files or folders, extracts a preview image for each item (first frame for GIFs, middle frame for videos, original for images), and assembles everything into one or more Word documents with captions and consistent page layout. The app shows progress and manages very large outputs by splitting at 400 MB per document.​​

Key features

Select files and folders: Pick individual media or entire directories; app recursively discovers supported formats.​

Broad format support: Images (JPG, PNG, BMP, TIFF, WEBP, GIF) and videos (MP4, MOV, AVI, MKV, WMV, FLV).​

Smart frame extraction: First frame from GIFs, mid-frame from videos; images are verified for integrity.​

Word gallery output: Adds file path label, then an inserted picture with fixed width for uniform look.​

Auto document splitting: Keeps each .docx under ~400 MB, creating part-named files automatically.​

Live progress: Determinate progress bar with UI kept responsive via a worker thread.​

Size feedback: Shows combined size of selected inputs for user awareness.​

How it works

Input selection: Use “Wybierz pliki” to choose media or “Wybierz folder” to crawl a directory tree; selected items are aggregated and deduplicated by path order.​

Validation: Images are opened with Pillow verify() to detect corrupt files; GIFs are converted to a PNG snapshot; videos are opened with OpenCV to grab the middle frame.​

Document layout: Creates a .docx with A4 page, 15 mm margins, and a heading; each media entry adds the source path then a 120 mm-wide image, with spacing after.​

Size control: After each addition the temporary document is saved to measure size; when ≥400 MB, the part is saved and a new document begins.​

Completion: The final part is saved with a “_partN.docx” suffix and a success message is shown.​

Supported formats

Images: .jpg, .jpeg, .png, .bmp, .tiff, .tif, .webp, .gif.​

Videos: .mp4, .mov, .avi, .mkv, .wmv, .flv.​

Installation

Requirements: Python 3.x, Tkinter (bundled with most Python installers), Pillow, OpenCV-Python, python-docx.​

Install packages: pip install pillow opencv-python python-docx.​

Launch: python your_script.py to open the GUI.​

User guide

Choose inputs:

Click “Wybierz pliki” to pick media files.​

Click “Wybierz folder” to add all supported media recursively.​

Set output file:

Click “Wybierz plik wyjściowy,” choose a .docx name; parts will be suffixed automatically.​

Start:

Click “Rozpocznij przetwarzanie” to begin; the progress bar will update while the UI remains responsive.​

Result:

The app saves one or more .docx files and displays a completion dialog with final path.​

Best practices

Prefer stable video codecs readable by OpenCV (e.g., MP4/H.264) to ensure frame extraction works reliably.​

For huge batches, split folders logically so each part stays below the 400 MB cap faster and reduces temp I/O.​

Keep sufficient disk space for temporary PNG frames used during extraction from GIFs and videos.​

Troubleshooting

Video frame not extracted:

Ensure the file opens in OpenCV; re-encode to MP4/H.264 if unsupported.​

Corrupt image:

Non-verified images are skipped; re-save the image or remove it from selection.​

Large output stalls:

Output size is checked frequently; wait for the part rollover or reduce batch size.​

Security and privacy

All processing is local; temporary images are written to the system temp directory and removed when no longer needed.​

Only paths and thumbnails are embedded in the .docx; no network activity occurs.​

Limitations

Audio is not embedded; only visual thumbnails are captured from videos.​

Highly compressed or exotic codecs may fail to decode; convert such files first.​

Credits

Built with Tkinter (GUI), Pillow (image verify/saving), OpenCV (video frame extraction), and python-docx (Word output).​
