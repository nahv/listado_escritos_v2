import webview
from backend import Api

if __name__ == '__main__':
    api = Api()
    window = webview.create_window('Listado de escritos', 'frontend/index.html', js_api=api, width=1200, height=933)
    webview.start()  # debug=False by default
