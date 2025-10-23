"""
Example usage of the vbaProjectCompiler builder module.

This example demonstrates how to use the builder helper functions
to create VBA projects from source files.
"""

from pathlib import Path
from vbaProjectCompiler.builder import build_from_directory, create_project_from_files


def example_1_build_from_directory():
    """
    Example 1: Build a VBA project from a structured directory.

    This is the simplest approach - just organize your VBA files
    in the expected directory structure and let the builder do the rest.
    """
    print("Example 1: Building from directory structure")
    print("=" * 50)

    # Assuming you have a directory structure like:
    # my_vba_project/
    #   Modules/
    #     Module1.bas
    #     Utils.bas
    #   Objects/
    #     ThisWorkbook.cls
    #     Sheet1.cls
    #   ClassModules/
    #     DataProcessor.cls

    try:
        project = build_from_directory("my_vba_project")
        print(f"Successfully created project with {len(project.modules)} modules:")
        for module in project.modules:
            print(f"  - {module.get_name()} ({module.type})")

        # Write the vbaProject.bin file using OleFile from MS-CFB package
        from ms_cfb import OleFile  # Note: OleFile is in the MS-CFB package

        ole_file = OleFile(project)
        ole_file.writeFile("vbaProject.bin")
        print("\nProject binary written to vbaProject.bin")

    except FileNotFoundError as e:
        print(f"Error: {e}")
    except ValueError as e:
        print(f"Error: {e}")

    print()


def example_2_build_from_files_dict():
    """
    Example 2: Build a VBA project from a custom file list.

    Use this approach when you have a non-standard directory structure
    or want more control over which files are included.
    """
    print("Example 2: Building from file dictionary")
    print("=" * 50)

    # Define your files explicitly
    files = {
        "modules": [
            "src/code/Module1.bas",
            "src/code/Utils.bas",
        ],
        "doc_modules": [
            "src/objects/ThisWorkbook.cls",
            "src/objects/Sheet1.cls",
        ],
        "class_modules": [
            "src/classes/DataProcessor.cls",
            "src/classes/FileHandler.cls",
        ],
        "forms": [
            "src/forms/UserForm1.frm",
        ],
    }

    # Note: This example won't actually work unless these files exist
    # It's just to demonstrate the API
    print("Would create project from:")
    for category, file_list in files.items():
        print(f"\n{category}:")
        for file in file_list:
            print(f"  - {file}")

    # Uncomment to actually create the project:
    # project = create_project_from_files(files)
    # print(f"\nProject created with {len(project.modules)} modules")

    print()


def example_3_customize_project():
    """
    Example 3: Build a project and customize it with references.

    This shows how you can use the builder and then add additional
    configuration like library references.
    """
    print("Example 3: Building and customizing a project")
    print("=" * 50)

    # First, build the basic project
    try:
        project = build_from_directory("my_vba_project")
        print(f"Created base project with {len(project.modules)} modules")

        # Now add custom configuration
        # For example, add library references:
        from vbaProjectCompiler.Models.Entities.referenceRecord import ReferenceRecord
        from vbaProjectCompiler.Models.Fields.libidReference import LibidReference
        from uuid import UUID

        # Add a reference to stdole (OLE Automation)
        libidRef = LibidReference(
            UUID("00020430-0000-0000-C000-000000000046"),
            "2.0",
            "0",
            "C:\\Windows\\System32\\stdole2.tlb",
            "OLE Automation",
        )
        oleReference = ReferenceRecord("cp1252", "stdole", libidRef)
        project.addReference(oleReference)
        print("Added OLE Automation reference")

        # Set project properties
        project.setProjectId("{12345678-1234-1234-1234-123456789012}")
        print("Set project ID")

        print("\nProject is now fully configured!")

    except Exception as e:
        print(f"Error: {e}")

    print()


def example_4_real_world_workflow():
    """
    Example 4: A real-world workflow for VBA project compilation.

    This demonstrates a complete workflow from extracting VBA from Excel
    to rebuilding the binary.
    """
    print("Example 4: Real-world workflow")
    print("=" * 50)

    print(
        """
    Typical workflow:
    
    1. Extract VBA from existing Excel file (using a tool like oletools or vba-extractor)
       - Excel file -> VBA source files
    
    2. Organize extracted files into the expected structure:
       my_project/
         Modules/
         ClassModules/
         Objects/
         Forms/
    
    3. Edit your VBA source files as needed
    
    4. Use the builder to compile back to vbaProject.bin:
    """
    )

    print("    from vbaProjectCompiler.builder import build_from_directory")
    print("    project = build_from_directory('my_project')")
    print("    ")
    print("    # Write the bin file using OleFile from MS-CFB package")
    print("    from ms_cfb import OleFile")
    print("    ole_file = OleFile(project)")
    print("    ole_file.writeFile('vbaProject.bin')")

    print(
        """
    5. Inject the vbaProject.bin back into your Excel file
       (using appropriate OLE tools)
    
    Benefits:
    - Version control for VBA code (use Git!)
    - Easier code review and collaboration
    - Automated testing and CI/CD for VBA projects
    - Separate editing of VBA from Excel files
    """
    )


if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("VBA Project Compiler - Usage Examples")
    print("=" * 50 + "\n")

    example_1_build_from_directory()
    example_2_build_from_files_dict()
    example_3_customize_project()
    example_4_real_world_workflow()

    print("\nFor more information, see README.md")
