from flask import Flask, request, jsonify
from flask_cors import CORS
import re

app = Flask(__name__)
CORS(app)

class VBAParser:
    def __init__(self, code):
        self.code = code
        self.functions = []
        self.subroutines = []
        self.variables = []
        self.comments = []
        self.loops = []
        self.conditionals = []
        self.error_handling = []
        self.classes = []

    def parse(self):
        self._extract_comments()
        self._extract_functions_and_subroutines()
        self._extract_variables()
        self._extract_loops()
        self._extract_conditionals()
        self._extract_error_handling()
        self._extract_classes()

    def _extract_comments(self):
        comment_pattern = re.compile(r"'(.*)")
        self.comments = comment_pattern.findall(self.code)

    def _extract_functions_and_subroutines(self):
        function_pattern = re.compile(r'Function\s+(\w+)\s*\((.*?)\)\s*As\s*(\w+)', re.IGNORECASE)
        subroutine_pattern = re.compile(r'Sub\s+(\w+)\s*\((.*?)\)', re.IGNORECASE)

        functions = function_pattern.findall(self.code)
        subroutines = subroutine_pattern.findall(self.code)

        for function in functions:
            name, params, return_type = function
            self.functions.append({
                'type': 'Function',
                'name': name,
                'params': params,
                'return_type': return_type,
                'body': self._extract_body(name)
            })

        for subroutine in subroutines:
            name, params = subroutine
            self.subroutines.append({
                'type': 'Sub',
                'name': name,
                'params': params,
                'body': self._extract_body(name)
            })

    def _extract_body(self, name):
        pattern = re.compile(rf'{name}\s*\(.*?\)(.*?)\s*End\s+(Sub|Function)', re.DOTALL | re.IGNORECASE)
        match = pattern.search(self.code)
        if match:
            return match.group(1).strip()
        return ''

    def _extract_variables(self):
        variable_pattern = re.compile(r'Dim\s+(\w+)\s+As\s+(\w+)', re.IGNORECASE)
        self.variables = variable_pattern.findall(self.code)

    def _extract_loops(self):
        loop_patterns = [
            re.compile(r'For\s+.*?Next', re.DOTALL | re.IGNORECASE),
            re.compile(r'Do\s+.*?Loop', re.DOTALL | re.IGNORECASE),
            re.compile(r'While\s+.*?Wend', re.DOTALL | re.IGNORECASE)
        ]
        for pattern in loop_patterns:
            self.loops.extend(pattern.findall(self.code))

    def _extract_conditionals(self):
        if_pattern = re.compile(r'If\s+.*?End\s+If', re.DOTALL | re.IGNORECASE)
        self.conditionals = if_pattern.findall(self.code)

    def _extract_error_handling(self):
        error_pattern = re.compile(r'On\s+Error\s+.*', re.IGNORECASE)
        self.error_handling = error_pattern.findall(self.code)

    def _extract_classes(self):
        class_pattern = re.compile(r'Class\s+(\w+)', re.IGNORECASE)
        self.classes = class_pattern.findall(self.code)

    def get_documentation(self):
        documentation = {
            'Comments': self.comments,
            'Variables': [{'Name': var[0], 'Type': var[1]} for var in self.variables],
            'Functions': [{'Name': func['name'], 'Params': func['params'], 'Returns': func['return_type'], 'Body': func['body']} for func in self.functions],
            'Subroutines': [{'Name': sub['name'], 'Params': sub['params'], 'Body': sub['body']} for sub in self.subroutines],
            'Loops': self.loops,
            'Conditionals': self.conditionals,
            'ErrorHandling': self.error_handling,
            'Classes': self.classes
        }
        return documentation

@app.route('/parse_vba', methods=['POST'])
def parse_vba():
    vba_code = request.json.get('code')
    if not vba_code:
        return jsonify({'error': 'No VBA code provided'}), 400

    parser = VBAParser(vba_code)
    parser.parse()
    documentation = parser.get_documentation()

    return jsonify({'documentation': documentation})

if __name__ == '__main__':
    app.run(debug=True)
