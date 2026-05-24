import os

def get_mime_type(file_path):
    # Extract the file extension from the file path
    _, file_extension = os.path.splitext(file_path)

    # Map common file extensions to MIME types
    mime_types = {
        '.txt': 'text/plain',
        '.pdf': 'application/pdf',
        '.jpg': 'image/jpeg',
        '.png': 'image/png',
        '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        # Add more file extensions and MIME types as needed
    }

    # Return the corresponding MIME type.
    return mime_types.get(file_extension.lower(), 'application/octet-stream')


if __name__ == "__main__":
# Test routine for get_mime_type function
    file_path = 'Präsentation1.pptx'  # Test some file path
    mime_type = get_mime_type(file_path)
    print(f'MIME type of {file_path}: {mime_type}')