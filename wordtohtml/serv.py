__author__ = 'alexander'
import os
import sys

from tornado.options import options, define, parse_command_line
import django.core.handlers.wsgi
import tornado.httpserver
import tornado.ioloop
import tornado.web
import tornado.wsgi
from django.core.wsgi import get_wsgi_application

_HERE = os.path.abspath(os.path.dirname(__file__))
sys.path.append(_HERE)
os.environ['DJANGO_SETTINGS_MODULE'] = "wordtohtml.settings"

def main(port):
    wsgi_app = tornado.wsgi.WSGIContainer(
        #django.core.handlers.wsgi.WSGIHandler()
        get_wsgi_application()
        )
    tornado_app = tornado.web.Application(
        [('.*', tornado.web.FallbackHandler, dict(fallback=wsgi_app)),
        ],debug=True)
    server = tornado.httpserver.HTTPServer(tornado_app)
    server.listen(port)
    tornado.ioloop.IOLoop.instance().start()


if __name__ == '__main__':
    main(int(sys.argv[1]))
