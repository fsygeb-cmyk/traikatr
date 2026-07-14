from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import os

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        response = {"message": "Hello from Python!"}
        self.wfile.write(json.dumps(response).encode())

def run():
    port = int(os.environ.get('PORT', 8000))
    server = HTTPServer(('', port), Handler)
    print(f"✅ Server running on port {port}")
    server.serve_forever()

if __name__ == "__main__":
    run()
