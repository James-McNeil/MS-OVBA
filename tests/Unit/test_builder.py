"""
Unit tests for the builder module.
"""
import pytest
import tempfile
import shutil
from pathlib import Path
from vbaProjectCompiler.builder import build_from_directory, create_project_from_files
from vbaProjectCompiler.vbaProject import VbaProject
from vbaProjectCompiler.Models.Entities.doc_module import DocModule
from vbaProjectCompiler.Models.Entities.std_module import StdModule


class TestBuildFromDirectory:
    """Tests for build_from_directory function."""
    
    def setup_method(self):
        """Create a temporary directory for each test."""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
    
    def teardown_method(self):
        """Clean up temporary directory after each test."""
        if Path(self.temp_dir).exists():
            shutil.rmtree(self.temp_dir)
    
    def create_sample_vba_file(self, filepath, content="Sub Test()\nEnd Sub\n"):
        """Helper to create a sample VBA file."""
        filepath.parent.mkdir(parents=True, exist_ok=True)
        filepath.write_text(content, encoding='utf-8')
    
    def test_nonexistent_directory_raises_error(self):
        """Test that FileNotFoundError is raised for non-existent directory."""
        with pytest.raises(FileNotFoundError, match="Source directory not found"):
            build_from_directory("nonexistent_dir")
    
    def test_file_instead_of_directory_raises_error(self):
        """Test that ValueError is raised when source is a file, not a directory."""
        test_file = self.temp_path / "test.txt"
        test_file.write_text("test")
        
        with pytest.raises(ValueError, match="Source path is not a directory"):
            build_from_directory(str(test_file))
    
    def test_empty_directory_raises_error(self):
        """Test that ValueError is raised when no VBA files are found."""
        with pytest.raises(ValueError, match="No VBA files found"):
            build_from_directory(self.temp_dir)
    
    def test_build_with_standard_modules(self):
        """Test building a project with standard .bas modules."""
        modules_dir = self.temp_path / "Modules"
        self.create_sample_vba_file(modules_dir / "Module1.bas")
        self.create_sample_vba_file(modules_dir / "Module2.bas")
        
        project = build_from_directory(self.temp_dir)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 2
        assert all(isinstance(m, StdModule) for m in project.modules)
        assert project.modules[0].get_name() == "Module1"
        assert project.modules[1].get_name() == "Module2"
    
    def test_build_with_class_modules(self):
        """Test building a project with class modules."""
        class_dir = self.temp_path / "ClassModules"
        self.create_sample_vba_file(class_dir / "MyClass.cls")
        
        project = build_from_directory(self.temp_dir)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 1
        assert isinstance(project.modules[0], StdModule)
        assert project.modules[0].get_name() == "MyClass"
    
    def test_build_with_document_modules(self):
        """Test building a project with document modules."""
        objects_dir = self.temp_path / "Objects"
        self.create_sample_vba_file(objects_dir / "ThisWorkbook.cls")
        self.create_sample_vba_file(objects_dir / "Sheet1.cls")
        
        project = build_from_directory(self.temp_dir)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 2
        assert all(isinstance(m, DocModule) for m in project.modules)
        assert project.modules[0].get_name() == "Sheet1"
        assert project.modules[1].get_name() == "ThisWorkbook"
    
    def test_build_with_form_modules(self):
        """Test building a project with form modules."""
        forms_dir = self.temp_path / "Forms"
        self.create_sample_vba_file(forms_dir / "UserForm1.frm")
        
        project = build_from_directory(self.temp_dir)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 1
        assert isinstance(project.modules[0], StdModule)
        assert project.modules[0].get_name() == "UserForm1"
    
    def test_build_with_mixed_modules(self):
        """Test building a project with all types of modules."""
        modules_dir = self.temp_path / "Modules"
        class_dir = self.temp_path / "ClassModules"
        objects_dir = self.temp_path / "Objects"
        forms_dir = self.temp_path / "Forms"
        
        self.create_sample_vba_file(modules_dir / "Module1.bas")
        self.create_sample_vba_file(class_dir / "MyClass.cls")
        self.create_sample_vba_file(objects_dir / "ThisWorkbook.cls")
        self.create_sample_vba_file(forms_dir / "UserForm1.frm")
        
        project = build_from_directory(self.temp_dir)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 4
    
    def test_modules_are_sorted_alphabetically(self):
        """Test that modules are added in alphabetical order."""
        modules_dir = self.temp_path / "Modules"
        self.create_sample_vba_file(modules_dir / "Zebra.bas")
        self.create_sample_vba_file(modules_dir / "Alpha.bas")
        self.create_sample_vba_file(modules_dir / "Beta.bas")
        
        project = build_from_directory(self.temp_dir)
        
        assert len(project.modules) == 3
        assert project.modules[0].get_name() == "Alpha"
        assert project.modules[1].get_name() == "Beta"
        assert project.modules[2].get_name() == "Zebra"
    
    def test_output_bin_parameter_is_accepted(self):
        """Test that output_bin parameter is accepted (for future use)."""
        modules_dir = self.temp_path / "Modules"
        self.create_sample_vba_file(modules_dir / "Module1.bas")
        
        # Should not raise an error
        project = build_from_directory(self.temp_dir, output_bin="custom.bin")
        assert isinstance(project, VbaProject)


class TestCreateProjectFromFiles:
    """Tests for create_project_from_files function."""
    
    def setup_method(self):
        """Create a temporary directory for each test."""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
    
    def teardown_method(self):
        """Clean up temporary directory after each test."""
        if Path(self.temp_dir).exists():
            shutil.rmtree(self.temp_dir)
    
    def create_sample_vba_file(self, filepath, content="Sub Test()\nEnd Sub\n"):
        """Helper to create a sample VBA file."""
        filepath.parent.mkdir(parents=True, exist_ok=True)
        filepath.write_text(content, encoding='utf-8')
    
    def test_empty_dict_creates_empty_project(self):
        """Test that an empty dict creates a project with no modules."""
        project = create_project_from_files({})
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 0
    
    def test_create_with_standard_modules(self):
        """Test creating a project with standard modules."""
        module1 = self.temp_path / "Module1.bas"
        module2 = self.temp_path / "Module2.bas"
        self.create_sample_vba_file(module1)
        self.create_sample_vba_file(module2)
        
        files = {'modules': [str(module1), str(module2)]}
        project = create_project_from_files(files)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 2
        assert all(isinstance(m, StdModule) for m in project.modules)
    
    def test_create_with_class_modules(self):
        """Test creating a project with class modules."""
        class1 = self.temp_path / "MyClass.cls"
        self.create_sample_vba_file(class1)
        
        files = {'class_modules': [str(class1)]}
        project = create_project_from_files(files)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 1
        assert isinstance(project.modules[0], StdModule)
    
    def test_create_with_doc_modules(self):
        """Test creating a project with document modules."""
        workbook = self.temp_path / "ThisWorkbook.cls"
        sheet = self.temp_path / "Sheet1.cls"
        self.create_sample_vba_file(workbook)
        self.create_sample_vba_file(sheet)
        
        files = {'doc_modules': [str(workbook), str(sheet)]}
        project = create_project_from_files(files)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 2
        assert all(isinstance(m, DocModule) for m in project.modules)
    
    def test_create_with_form_modules(self):
        """Test creating a project with form modules."""
        form = self.temp_path / "UserForm1.frm"
        self.create_sample_vba_file(form)
        
        files = {'forms': [str(form)]}
        project = create_project_from_files(files)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 1
        assert isinstance(project.modules[0], StdModule)
    
    def test_create_with_mixed_modules(self):
        """Test creating a project with all types of modules."""
        module = self.temp_path / "Module1.bas"
        class_mod = self.temp_path / "MyClass.cls"
        doc_mod = self.temp_path / "ThisWorkbook.cls"
        form = self.temp_path / "UserForm1.frm"
        
        self.create_sample_vba_file(module)
        self.create_sample_vba_file(class_mod)
        self.create_sample_vba_file(doc_mod)
        self.create_sample_vba_file(form)
        
        files = {
            'modules': [str(module)],
            'class_modules': [str(class_mod)],
            'doc_modules': [str(doc_mod)],
            'forms': [str(form)]
        }
        project = create_project_from_files(files)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 4
    
    def test_missing_keys_dont_raise_error(self):
        """Test that missing dictionary keys don't cause errors."""
        module = self.temp_path / "Module1.bas"
        self.create_sample_vba_file(module)
        
        # Only provide 'modules' key, others are missing
        files = {'modules': [str(module)]}
        project = create_project_from_files(files)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 1
    
    def test_empty_lists_dont_raise_error(self):
        """Test that empty lists in the dict don't cause errors."""
        files = {
            'modules': [],
            'class_modules': [],
            'doc_modules': [],
            'forms': []
        }
        project = create_project_from_files(files)
        
        assert isinstance(project, VbaProject)
        assert len(project.modules) == 0
