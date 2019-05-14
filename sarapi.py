from flask import Flask, request, Response
import json

app = Flask(__name__)

@app.route("/get", methods=['GET'])
def get_network_info():
    mydict = {}
    mydict["routing_protocol"] = ""
    mydict["router1_port_channeling_type"] = ""
    mydict["router1_mtu"] = ""
    mydict["router1_target_network"] = ""
    mydict["Client_IPpool_CIDR"] = ""
    mydict["OCC1_handoff01_connected_router"] = ""
    mydict["OCC1_handoff01_subnet"] = ""
    mydict["OCC1_handoff01_upstreamIP"] = ""
    mydict["OCC1_handoff01_TorIP"] = ""
    return Response(json.dumps(mydict, indent=4), status=200, mimetype='application/json')


@app.route("/put", methods=['PUT'])
def check_network_info():
    resp = {}
    print(request.data)
    if not request.json:
        resp["error"] = 'no json payload'
        return Response(json.dumps(resp, indent=4), status=404)
    else:
        resp["Ok"] = request.json["param"]
        return Response(json.dumps(resp, indent=4), status=201, mimetype='application/json') */

@app.route("/", methods=['GET','POST','PUT'])
def check_network_info():
    resp = {}
    print(request.__dict__)
    if not request.json:
        resp["error"] = 'no json payload'
        return Response(json.dumps(resp, indent=4), status=404)
    else:
        resp["Ok"] = request.json["param"]
        return Response(json.dumps(resp, indent=4), status=201, mimetype='application/json')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)