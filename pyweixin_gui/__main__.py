try:
    from .app import main
except ImportError:
    # Allow direct execution such as `python pyweixin_gui/__main__.py`
    # and reduce launcher fragility in some packaging/runtime contexts.
    from pyweixin_gui.app import main


if __name__ == "__main__":
    raise SystemExit(main())
