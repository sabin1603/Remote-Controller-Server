import os

def find_pptx_files(root_dir):
    """
    Recursively find all .pptx files starting from root_dir.
    """
    pptx_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.pptx'):
                pptx_files.append(os.path.join(dirpath, filename))
    return pptx_files
