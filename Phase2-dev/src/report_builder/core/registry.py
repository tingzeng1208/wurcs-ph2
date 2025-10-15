# workbook_builder/core/registry.py
from typing import Callable
from registry_store import _REGISTRY  # <— single shared dict

def register(name: str):
    """Decorator to register a worksheet builder under a name."""
    def deco(fn):
        _REGISTRY[name] = fn
        return fn
    return deco

def get(name: str):
    """Fetch a registered worksheet builder by name."""
    return _REGISTRY[name]

def all_names() -> list[str]:
    return sorted(_REGISTRY.keys())

def load_all_plugins(package_path):
    """
    Import every submodule in a package directory
    so that any @register decorators run.
    """
    import importlib, pkgutil, os, sys
    
    # If package_path is a string, convert it to a proper module path
    if isinstance(package_path, str):
        # Add the parent directory to sys.path if not already there
        parent_path = os.path.dirname(package_path)
        if parent_path not in sys.path:
            sys.path.insert(0, parent_path)
        
        # Get the module name from the path
        module_name = os.path.basename(package_path)
        
        # Try to import the package
        try:
            package = importlib.import_module(module_name)
        except ImportError:
            print(f"Warning: Could not import package '{module_name}' from path '{package_path}'")
            return
    else:
        package = package_path
    
    # Load all submodules
    try:
        for _, modname, _ in pkgutil.iter_modules(package.__path__):
            importlib.import_module(f"{package.__name__}.{modname}")
    except AttributeError:
        print(f"Warning: Package {package} does not have __path__ attribute")
