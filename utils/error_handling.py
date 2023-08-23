class CustomError(Exception):
    """Base class for custom exceptions."""
    pass

class InvalidDataError(CustomError):
    """Raised when data extracted from Excel is invalid."""
    pass

class TemplateNotFoundError(CustomError):
    """Raised when a template file is not found."""
    pass

class PlaceholderNotFoundError(CustomError):
    """Raised when a placeholder is not found in a Word document."""
    pass
