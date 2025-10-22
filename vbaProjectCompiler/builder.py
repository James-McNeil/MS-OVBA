"""
Helper utilities for building VBA projects from source directories.

This module provides convenient functions to build VBA projects from
a directory structure containing VBA source files.
"""
from pathlib import Path
from vbaProjectCompiler.vbaProject import VbaProject
from vbaProjectCompiler.Models.Entities.doc_module import DocModule
from vbaProjectCompiler.Models.Entities.std_module import StdModule


def build_from_directory(source_dir, output_bin="vbaProject.bin"):
    """
    Build vbaProject.bin from a directory containing VBA source files.
    
    Expected directory structure:
        source_dir/
            Modules/        - Standard modules (.bas)
            ClassModules/   - Class modules (.cls)
            Objects/        - Document modules (.cls) like ThisWorkbook, Sheet1
            Forms/          - Form modules (.frm with optional .frx)
    
    Args:
        source_dir: Path to directory containing VBA source files
        output_bin: Path where vbaProject.bin should be written
        
    Returns:
        VbaProject: The configured VBA project object
        
    Raises:
        FileNotFoundError: If source_dir doesn't exist
        ValueError: If no VBA files are found in the directory structure
        
    Example:
        >>> from vbaProjectCompiler.builder import build_from_directory
        >>> project = build_from_directory("my_vba_project")
        >>> # Once OleFile is implemented:
        >>> # from vbaProjectCompiler.ole_file import OleFile
        >>> # ole_file = OleFile(project)
        >>> # ole_file.writeFile(output_bin)
        
    Note:
        This function currently returns a VbaProject object. Once the OleFile
        class is implemented, this function will also write the vbaProject.bin
        file directly.
    """
    source_path = Path(source_dir)
    
    if not source_path.exists():
        raise FileNotFoundError(f"Source directory not found: {source_dir}")
    
    if not source_path.is_dir():
        raise ValueError(f"Source path is not a directory: {source_dir}")
    
    project = VbaProject()
    module_count = 0
    
    # Add Standard Modules (.bas files)
    modules_dir = source_path / "Modules"
    if modules_dir.exists() and modules_dir.is_dir():
        for bas_file in sorted(modules_dir.glob("*.bas")):
            module = StdModule(bas_file.stem)
            module.add_file(str(bas_file))
            project.addModule(module)
            module_count += 1
    
    # Add Class Modules (.cls files)
    class_modules_dir = source_path / "ClassModules"
    if class_modules_dir.exists() and class_modules_dir.is_dir():
        for cls_file in sorted(class_modules_dir.glob("*.cls")):
            module = StdModule(cls_file.stem)
            module.add_file(str(cls_file))
            project.addModule(module)
            module_count += 1
    
    # Add Document Modules (.cls files like ThisWorkbook, Sheet1, etc.)
    objects_dir = source_path / "Objects"
    if objects_dir.exists() and objects_dir.is_dir():
        for cls_file in sorted(objects_dir.glob("*.cls")):
            module = DocModule(cls_file.stem)
            module.add_file(str(cls_file))
            project.addModule(module)
            module_count += 1
    
    # Add Form Modules (.frm files)
    forms_dir = source_path / "Forms"
    if forms_dir.exists() and forms_dir.is_dir():
        for frm_file in sorted(forms_dir.glob("*.frm")):
            module = StdModule(frm_file.stem)
            module.add_file(str(frm_file))
            project.addModule(module)
            module_count += 1
    
    if module_count == 0:
        raise ValueError(
            f"No VBA files found in {source_dir}. "
            "Expected subdirectories: Modules/, ClassModules/, Objects/, Forms/"
        )
    
    # TODO: Once OleFile is implemented, uncomment the following:
    # from vbaProjectCompiler.ole_file import OleFile
    # ole_file = OleFile(project)
    # ole_file.writeFile(output_bin)
    # return Path(output_bin)
    
    return project


def create_project_from_files(files_dict):
    """
    Create a VBA project from a dictionary of files.
    
    Args:
        files_dict: Dictionary with keys 'modules', 'class_modules', 
                   'doc_modules', 'forms' mapping to lists of file paths
                   
    Returns:
        VbaProject: The configured VBA project object
        
    Example:
        >>> files = {
        ...     'modules': ['Module1.bas', 'Module2.bas'],
        ...     'doc_modules': ['ThisWorkbook.cls', 'Sheet1.cls'],
        ...     'class_modules': ['MyClass.cls']
        ... }
        >>> project = create_project_from_files(files)
    """
    project = VbaProject()
    
    # Add standard modules
    for module_path in files_dict.get('modules', []):
        path = Path(module_path)
        module = StdModule(path.stem)
        module.add_file(str(path))
        project.addModule(module)
    
    # Add class modules
    for module_path in files_dict.get('class_modules', []):
        path = Path(module_path)
        module = StdModule(path.stem)
        module.add_file(str(path))
        project.addModule(module)
    
    # Add document modules
    for module_path in files_dict.get('doc_modules', []):
        path = Path(module_path)
        module = DocModule(path.stem)
        module.add_file(str(path))
        project.addModule(module)
    
    # Add form modules
    for module_path in files_dict.get('forms', []):
        path = Path(module_path)
        module = StdModule(path.stem)
        module.add_file(str(path))
        project.addModule(module)
    
    return project
