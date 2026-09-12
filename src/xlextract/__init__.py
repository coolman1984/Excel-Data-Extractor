"""The xlextract package and its temporary T-001 version entry point."""

__version__ = "0.1.0"


def main() -> None:
    """Print the package version until the CLI command package exists."""
    import sys

    sys.stdout.write(f"xlextract {__version__}\n")
