from __future__ import annotations


class ReportGeneratorError(Exception):
    """Base exception for user-facing conversion failures."""


class MissingInputFileError(ReportGeneratorError):
    """Raised when the selected source file does not exist."""


class MissingWorksheetError(ReportGeneratorError):
    """Raised when the source workbook does not contain the expected sheet."""


class MissingRequiredColumnsError(ReportGeneratorError):
    """Raised when one or more required source columns are absent."""


class InvalidDateError(ReportGeneratorError):
    """Raised when a source row contains an invalid Job Date."""


class InvalidAmountError(ReportGeneratorError):
    """Raised when a source row contains an invalid Loc Amt."""


class OutputPermissionError(ReportGeneratorError):
    """Raised when the output file cannot be written."""
