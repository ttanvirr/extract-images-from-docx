
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import shutil
import posixpath


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def extract_images():
    file_path = Path(input("Enter Word document path: ").strip().strip('"'))

    if not file_path.is_file():
        print("Error: File not found.")
        return

    if file_path.suffix.lower() != ".docx":
        print("Error: Please provide a .docx file.")
        return

    folder_name = file_path.stem
    output_folder = file_path.parent / f"{folder_name}_images"
    output_folder.mkdir(exist_ok=True)

    try:
        with ZipFile(file_path, "r") as docx:
            relationships = ET.fromstring(
                docx.read("word/_rels/document.xml.rels")
            )

            rel_map = {
                rel.attrib["Id"]: rel.attrib["Target"]
                for rel in relationships
                if rel.attrib.get("TargetMode") != "External"
            }

            document = ET.fromstring(docx.read("word/document.xml"))

            image_ids = []

            # Read images in the order they appear in the document.
            for element in document.iter():
                if element.tag.endswith("}blip"):
                    rid = element.attrib.get(f"{{{NS['r']}}}embed")
                    if rid:
                        image_ids.append(rid)

                elif element.tag.endswith("}imagedata"):
                    rid = element.attrib.get(f"{{{NS['r']}}}id")
                    if rid:
                        image_ids.append(rid)

            if not image_ids:
                print("No embedded images found in the document body.")
                return

            # Resolve each image reference to its media file.
            media_files = []
            for rid in image_ids:
                target = rel_map.get(rid)
                if not target:
                    continue

                media_path = posixpath.normpath(
                    posixpath.join("word", target)
                )

                if media_path in docx.namelist():
                    media_files.append(media_path)

            # Extract sequentially, waiting for each file write to finish.
            for index, media_path in enumerate(media_files, start=1):
                extension = Path(media_path).suffix
                output_file = (
                    output_folder / f"{index}-{folder_name}{extension}"
                )

                with docx.open(media_path) as source:
                    with output_file.open("wb") as destination:
                        shutil.copyfileobj(source, destination)

                print(f"Saved {index}/{len(media_files)}: {output_file.name}")

        print(f"\nFinished! Extracted {len(media_files)} image references.")
        print(f"Output folder: {output_folder}")

    except Exception as error:
        print(f"Error extracting images: {error}")


if __name__ == "__main__":
    extract_images()
