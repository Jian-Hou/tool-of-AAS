"""Compatibility entry point; failures now return a nonzero exit code."""
from qa import main
if __name__ == '__main__':
    raise SystemExit(main())
