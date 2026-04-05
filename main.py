
from app.window import QuickSnippetWindow, add_menu_bar
from app.startup import run_app, ensure_default_data

if __name__ == "__main__":
    run_app(QuickSnippetWindow, add_menu_bar, ensure_default_data)
