import webview
import os
import sys
from backend import Api


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


if __name__ == '__main__':
    api = Api()
    html_path = resource_path("frontend/index.html")
    window = webview.create_window(
        'Listado de escritos',
        html_path,
        js_api=api,
        width=1200,
        height=933
    )
    webview.start()