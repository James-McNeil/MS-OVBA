[![Coverage Status](https://coveralls.io/repos/github/Beakerboy/vbaProject-Compiler/badge.svg?branch=main)](https://coveralls.io/github/Beakerboy/vbaProject-Compiler?branch=main)
# vbaProject-Compiler
Create a vbaProject.bin file from VBA source files.

## Installation

```bash
pip install -r requirements.txt
```

For development (includes testing tools):

```bash
pip install -r requirements_dev.txt
pip install -e .
```

**Note:** The `OleFile` class is provided by the [MS-CFB](https://github.com/Beakerboy/MS-CFB) package (dev branch), which handles the OLE/CFB file format. This is automatically installed via requirements.txt.

## Quick Start with Builder Helper

The easiest way to build a VBA project is using the `builder` module, which automatically discovers and loads VBA files from a structured directory.

For detailed documentation, see [Builder Helper Feature Documentation](docs/FEATURE_BUILDER.md).

```python
from vbaProjectCompiler.builder import build_from_directory

# Build a project from a directory structure
project = build_from_directory("my_vba_project")

# Write the vbaProject.bin file using OleFile from MS-CFB package
from ms_cfb import OleFile  # Note: OleFile is in the MS-CFB package
ole_file = OleFile(project)
ole_file.writeFile("vbaProject.bin")
```

### Expected Directory Structure

The builder expects your VBA source files to be organized in the following structure:

```
my_vba_project/
├── Modules/        # Standard modules (.bas files)
│   ├── Module1.bas
│   └── Module2.bas
├── ClassModules/   # Class modules (.cls files)
│   └── MyClass.cls
├── Objects/        # Document modules (.cls files) like ThisWorkbook, Sheet1
│   ├── ThisWorkbook.cls
│   └── Sheet1.cls
└── Forms/          # Form modules (.frm files)
    └── UserForm1.frm
```

### Alternative: Build from File Dictionary

If you prefer more control or have a different file structure, you can use `create_project_from_files`:

```python
from vbaProjectCompiler.builder import create_project_from_files

files = {
    'modules': ['path/to/Module1.bas', 'path/to/Module2.bas'],
    'doc_modules': ['path/to/ThisWorkbook.cls', 'path/to/Sheet1.cls'],
    'class_modules': ['path/to/MyClass.cls'],
    'forms': ['path/to/UserForm1.frm']
}

project = create_project_from_files(files)
```

## VBAProject Class

The vbaProject class contains all the data and metadata that will be used to create the OLE container.

**Note:** The `OleFile` class is provided by the separate [MS-CFB](https://github.com/Beakerboy/MS-CFB) package.

```python
from vbaProjectCompiler.vbaProject import VbaProject
from ms_cfb import OleFile  # OleFile is in the MS-CFB package


project = VbaProject()
thisWorkbook = DocModule("ThisWorkbook")
thisWorkbook.addFile(path)
project.addModule(thisWorkbook)

ole_file = OleFile(project)
ole_file.writeFile(".")
```

The VbaProject class has many layers of customization available. Forexample a librry referenece can be added to the project.

```python
codePage = 0x04E4
codePageName = "cp" + str(codePage)
libidRef = LibidReference(
    "windows",
    "{00020430-0000-0000-C000-000000000046}",
    "2.0",
    "0",
    "C:\\Windows\\System32\\stdole2.tlb",
    "OLE Automation"
)
oleReference = ReferenceRecord(codePageName, "stdole", libidRef)
project.addReference(oleReference)
```

## oleFile Class

**Note:** The `OleFile` class is implemented in the separate [MS-CFB](https://github.com/Beakerboy/MS-CFB) package, not in MS-OVBA.

Users should not have to interact with the oleFile class directly. Its job is to extract the data from the vbaProject and turn it into a valid file. This includes deciding which data stream appears where, and applying different views to the models to save the data in the correct formats.

The oleFIle has two parts, a header and a FAT Sector Chain. This FAT chain stores multiple streams of data:
* Fat Chain Stream
* Directory Stream
* Minifat Chain Stream
* Minifat Data Stream
* Fat Data Stream

These are all different views of data from the following Models

* fatChain
* minifatChain
* directoryStream
