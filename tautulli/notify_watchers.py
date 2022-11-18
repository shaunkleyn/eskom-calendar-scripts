import sys
import getopt
try:
    from urllib.request import Request, urlopen  # Python 3
except ImportError:
    from urllib2 import Request, urlopen  # Python 2
import json
import sys
import configparser
from datetime import datetime, timedelta
from string import Template
import logging

def main(argv):
    logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s")
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    root.addHandler(handler)

    default_endtime = hour_rounder(datetime.now() + timedelta(hours=2)).strftime("%H:%M")
    arg_endtime = ""
    arg_starttime = ""
    arg_input = ""
    arg_output = ""
    arg_user = ""
    arg_help = "{0} -e <end-time> -s <start-time> -l <log-file>".format(argv[0])
    
    try:
        opts, args = getopt.getopt(argv[1:], "he:s:i:u:o:", ["help", "endtime=", "starttime=","input=", 
        "user=", "output="])
    except:
        print(arg_help)
        sys.exit(2)
    
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(arg_help)  # print the help message
            sys.exit(2)
        elif opt in ("-e", "--endtime"):
            arg_endtime = arg
        elif opt in ("-s", "--starttime"):
            arg_starttime = arg
        elif opt in ("-i", "--input"):
            arg_input = arg
        elif opt in ("-u", "--user"):
            arg_user = arg
        elif opt in ("-o", "--output"):
            arg_output = arg

    if arg_endtime == "":
        arg_endtime = default_endtime
        logging.debug('No end time provided.End Time defaulted to current time + 2 hours rounded to the closest hour')
        logging.debug('End Time defaulted to {0} (current time + 2 hours rounded to the closest hour)'.format(default_endtime))

    config = configparser.ConfigParser()
    config.sections()
    config.read('configuration.ini')

    tautulli = config['TAUTULLI']
    tautulli['url']
    tautulli['port']
    tautulli['api_key']

    tv_message = Template(config['DEFAULT']['tvshow_termination_message'])
    movie_message = Template(config['DEFAULT']['movie_termination_message'])

    protocol = get_protocol(tautulli['port'])
    logging.debug('Using protocol ' + protocol)

    req = Request(protocol + '://' + tautulli['url'] + ':' + tautulli['port'] + '/api/v2?apikey=' + tautulli['api_key'] + '&cmd=get_activity')
    req.add_header('User-Agent', 'Mozilla/5.0')
    logging.debug('Sending request ' + req.full_url)

    content = urlopen(req).read()
    json_data = json.loads(content)
    streams = json_data["response"]["data"]["sessions"]
    logging.debug('{0} streams currently playing'.format(len(streams)))

    for i in streams:
        termination_reason = ''
        if i["media_type"] == "episode":
            termination_reason = tv_message.substitute(season_number=i["parent_media_index"].zfill(2),
                episode_number=i["media_index"].zfill(2),
                grandparent_title=i["grandparent_title"],
                endtime=arg_endtime)
        elif i["media_type"] == "movie":
            termination_reason = movie_message.substitute(title=i["title"],
                endtime=arg_endtime)
        logging.debug('Sending termination message to ' + i["user"] + ' : ' + termination_reason)

def hour_rounder(t):
    # Rounds to nearest hour by adding a timedelta hour if minute >= 30
    return (t.replace(second=0, microsecond=0, minute=0, hour=t.hour)
               +timedelta(hours=t.minute//30))

def get_protocol(port):
    protocol = ''
    port_number = int(port)
    if port_number == '':
        protocol = ''
    elif port_number == 80:
        protocol = 'http'
    elif port_number == 443:
        protocol = 'https'
    return protocol

if __name__ == "__main__":
    main(sys.argv)

