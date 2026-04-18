import sys
import os
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from app import app, init_db
    init_db()
except Exception as e:
    import flask
    _err = traceback.format_exc()
    print(f"[STARTUP ERROR] {_err}")

    # 에러 내용을 /error 페이지로 확인 가능하게 노출 (디버깅용)
    _app = flask.Flask(__name__)
    @_app.route('/', defaults={'path': ''})
    @_app.route('/<path:path>')
    def error_page(path):
        return flask.Response(
            f"<pre>시작 오류:\n{_err}</pre>",
            status=500, mimetype='text/html'
        )
    app = _app
