import argparse

from app.web.server import run_app


def main():
    parser = argparse.ArgumentParser(description="InsightFlow — relatórios de frota")
    parser.add_argument(
        "--dev",
        action="store_true",
        help="Inicia o Vite com hot reload (http://127.0.0.1:5173). A API continua no Flask.",
    )
    parser.add_argument(
        "--rebuild",
        action="store_true",
        help="Força npm install e npm run build antes de subir (modo normal).",
    )
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--no-browser", action="store_true", help="Nao abre o navegador automaticamente.")
    args = parser.parse_args()

    run_app(
        host=args.host,
        port=args.port,
        open_browser=not args.no_browser,
        dev_mode=args.dev,
        rebuild_frontend=args.rebuild,
    )


if __name__ == "__main__":
    main()
