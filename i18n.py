import json
import os

_DIR = os.path.join(os.path.dirname(__file__), 'i18n')
_CACHE = {}
SUPPORTED = ['ko', 'en', 'ja', 'zh_hans', 'zh_hant']
LANG_LABELS = {
    'ko':      '한국어',
    'en':      'English',
    'ja':      '日本語',
    'zh_hans': '简体中文',
    'zh_hant': '繁體中文',
}

def _load(lang):
    if lang not in _CACHE:
        path = os.path.join(_DIR, f'{lang}.json')
        try:
            with open(path, encoding='utf-8') as f:
                _CACHE[lang] = json.load(f)
        except FileNotFoundError:
            _CACHE[lang] = {}
    return _CACHE[lang]

def make_T(lang):
    trans = _load(lang)
    ko    = _load('ko')
    def T(key):
        return trans.get(key) or ko.get(key, key)
    return T
