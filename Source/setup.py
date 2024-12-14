from cx_Freeze import setup, Executable
import sys

# Dependencies
build_exe_options = {
    "packages": ["tkinter", "json", "logging", "pathlib", "typing", "dataclasses", 
                "reportlab.lib", "reportlab.platypus", "reportlab.pdfbase"],
    "include_files": [
        # Add any additional files your program needs here
    ],
    "excludes": ["test", "unittest"],
}

# Base for GUI applications
base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="Character Card Extractor",
    version="1.0",
    description="Extract and format character card data",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "card_extractor.py",  # your main python file
            base=base,
            target_name="CharacterCardExtractor.exe",
            icon="card_extractor.ico",  # added icon file path here
        )
    ]
)