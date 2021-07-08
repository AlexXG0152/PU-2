# -*- coding: utf-8 -*- 
from bottle import route, run, request, put, error
import os
import pandas as pd
from pathlib import Path
import json
import webbrowser
from threading import Timer

# template HTML page for simplicity and minimalize script. Include CSS ans JS.
template = """<!DOCTYPE html>
<html>
  <meta charset="utf-8">
  <head>
  <!--
    DO NOT SIMPLY COPY THOSE LINES. Download the JS and CSS files from the
    latest release (https://github.com/enyo/dropzone/releases/latest), and
    host them yourself!
  -->
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
  <script src="https://rawgit.com/enyo/dropzone/master/dist/dropzone.js"></script>
  <link rel="stylesheet" href="https://rawgit.com/enyo/dropzone/master/dist/dropzone.css">
  </head>
  <body>
  <!-- Change /upload-target to your upload address -->
  <form action="/upload" method="post" class="dropzone" enctype="multipart/form-data" name="upload" id="upload"></form>

  <button type="button" onclick="onSubmit('/convert')">Convert</button>

  <p id="result"></p>
  <a href="someurl">
    <p id="response"></p>
  </a>

  <script>
    function onSubmit(value) {
      var myRequest = new Request("/convert");
      var myInit = { method: "PUT" };
      fetch(myRequest, myInit).then(response => response.text())
        .then((response) => {
            $("#response").text(response)
            $("a").prop("href", (response))
            $("#result").text("Your result .xlsx file in this folder:")
        });
    }
  </script>

  <script>
  Dropzone.options.upload = {
        acceptedFiles:'.json, .txt'       
    };
  </script>  
  
  </body>
</html>
"""


file = [0] # save uploaded filename


@error(500)
def error500(error):
    return 'CHECK YOUR INPUT FILE!'


def exception_handler(func):
    """ Exception handler for user input in functions """
    def inner_function(*args, **kwargs):
        while True:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                return "ERROR!"
    return inner_function


@route('/')
def hello():
    """Return template page"""
    return template


@route('/upload', method='POST', name='submit')
def upload():
    """
    Upload func accept two file type and save this files in folder.
    If folder is not exist - func will create PU2 folder.    
    """
    upload = request.files.file
    name, ext = os.path.splitext(upload.filename)
    if ext not in ('.json', '.txt'):
        return 'File extension not allowed.'

    save_path = '/PU2/'

    if not os.path.isdir(save_path):
        os.mkdir(save_path)
    
    upload.save(save_path, overwrite=True)
    file[0] = os.path.abspath(upload.filename)

    return file[0]


@exception_handler
@put('/convert')
def convert():
    """Open uploaded file on server side and processing file with Pandas."""
    data = open_file()
    process(data)
    return f'file:///{os.path.dirname(file[0])}\\result.xlsx'


def flatten_json(nested_json: dict, exclude: list=[''], sep: str='_') -> dict:
    """
    Flatten a list of nested dicts.
    """
    out = dict()
    def flatten(x: (list, dict, str), name: str='', exclude=exclude):
        if type(x) is dict:
            for a in x:
                if a not in exclude:
                    flatten(x[a], f'{name}{a}{sep}')
        elif type(x) is list:
            i = 0
            for a in x:
                flatten(a, f'{name}{i}{sep}')
                i += 1
        else:
            out[name[:-1]] = x

    flatten(nested_json)
    return out


def open_file():
    """Open uploaded file on server side."""
    p = Path(file[0])
    with p.open('r', encoding='utf-8') as f: #mbcs (one more encoding type)
        data = json.loads(f.read())

    return data


def process(data):
    """
    Func for processing open file with Pandas.
    Flattened JSON type data from opened file, remove trash symbols from columns names.
    Rename colums names to understandable Russian names.
    """
    df = pd.json_normalize(data)
    df = pd.DataFrame([flatten_json(x) for x in data['data']])
    df.rename(columns=lambda x: x.rpartition('_')[-1] if '_' in x else x, inplace=True)
    
    con_names = {
        "ils": "страховой номер",
        "fzl": "фамилия",
        "izl": "собственное имя",
        "ozl": "отчество (если таковое имеется)",
        "dfr1": "дата приема на работу",
        "dpr11": "дата приказа о приеме на работу",
        "npr11": "номер приказа о приеме на работу",
        "dto1": "дата увольнения с работы",
        "dpr12": "дата приказа об увольнении с работы",
        "npr12": "номер приказа об увольнении с работы",
        "koduv1": "код основания увольнения",
        "prof2": "код должности служащего, профессии рабочего",
        "nameprf": "наименование должности служащего, профессии рабочего",
        "podr": "наименование структурного подразделения",
        "sovm": "код работы по совместительству",
        "dfr3": "дата приема",
        "dpr31": "дата приказа о назначении на должность служащего (профессию рабочего)",
        "npr31": "номер приказа о назначении на должность служащего (профессию рабочего)",
        "kod": "код вида трудового договора",
        "dto3": "дата увольнения",
        "dpr32": "дата приказа об увольнении с работы",
        "npr32": "номер приказа об увольнении с работы",
        "koduv3": "код основания увольнения",
        "dfr22": "дата присвоения квалификационной категории",
        "dpr22": "дата приказа о присвоении квалификационной категории",
        "npr22": "номер приказа о присвоении квалификационной категории",
        "rzrd": "разряд",
        "ktgr": "квалификационная категория",
        "kls": "класс",
        "klsgos": "класс государственного служащего"
        }

    for k, v in con_names.items():
        df.loc[-1, k] = v

    df.index = df.index + 1
    df = df.sort_index()
    df.to_excel('result.xlsx', index=False, encoding='utf-8')
    
    return df


def open_browser():
    """Automate open page in browser after start script"""
    webbrowser.open_new('http://127.0.0.1:8080/')


if __name__ == '__main__':
    Timer(1, open_browser).start()
    run(port=8080, debug=True)
