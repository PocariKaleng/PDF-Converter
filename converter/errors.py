class ConversionError(Exception):
    """A useful error that can be shown to the user."""

    def __init__(self, message: str, status: int = 422):
        super().__init__(message)
        self.status = status
