import flask
import os
from openpyxl import Workbook
from openpyxl.writer.excel import save_virtual_workbook
from pdq39xls import *

app = flask.Flask(__name__)
app.config["DEBUG"] = True

@app.route('/pdq39/data', methods=['GET'])
def download():
    ip = os.getenv('HOST_IP')
    port = os.getenv('CUDUI_PORT')
    cudui = f'http://{ip}:{port}'

    query_parameters = flask.request.args
    pid = query_parameters.get('pid')
    wb = pdq39xls(pid)
    return flask.Response(
        save_virtual_workbook(wb),
        headers={
            'Content-Disposition': 'attachment; filename=pdq39data.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Access-Control-Allow-Origin' : cudui
        }
    )

app.run()