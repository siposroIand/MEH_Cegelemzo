import os

from waitress import serve

from web.app import create_app


def main():
    app = create_app()
    try:
        serve(app, host=os.environ.get("MEH_HOST", "127.0.0.1"),
              port=int(os.environ.get("MEH_PORT", "8080")), threads=4)
    finally:
        app.extensions["workspace"].close()


if __name__ == "__main__":
    main()
