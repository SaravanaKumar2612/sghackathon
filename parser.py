import re
import openai

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
        self.errors = []

    def parse(self):
        self._extract_comments()
        self._extract_functions_and_subroutines()
        self._extract_variables()
        self._extract_loops()
        self._extract_conditionals()
        self._extract_error_handling()
        self._extract_classes()
        self._detect_errors()

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

    def _detect_errors(self):
        # Check for undeclared variables
        declared_variables = {var[0] for var in self.variables}
        used_variables = set(re.findall(r'\b(\w+)\b', self.code)) - declared_variables
        if used_variables:
            self.errors.append(f"Undeclared variables used: {', '.join(used_variables)}")

        # Check for unmatched loops and conditionals
        if len(re.findall(r'\bFor\b', self.code)) != len(re.findall(r'\bNext\b', self.code)):
            self.errors.append("Unmatched For/Next statements found.")
        if len(re.findall(r'\bIf\b', self.code)) != len(re.findall(r'\bEnd If\b', self.code)):
            self.errors.append("Unmatched If/End If statements found.")
        if len(re.findall(r'\bDo\b', self.code)) != len(re.findall(r'\bLoop\b', self.code)):
            self.errors.append("Unmatched Do/Loop statements found.")
        if len(re.findall(r'\bWhile\b', self.code)) != len(re.findall(r'\bWend\b', self.code)):
            self.errors.append("Unmatched While/Wend statements found.")

        # Check for unused variables
        for var in declared_variables:
            if var not in re.findall(r'\b' + re.escape(var) + r'\b', self.code):
                self.errors.append(f"Declared variable '{var}' is not used.")

    def get_human_readable_documentation(self, openai_api_key):
        documentation = self._generate_documentation_text()
        chatgpt_response = self._generate_chatgpt_response(documentation, openai_api_key)
        return chatgpt_response

    def _generate_documentation_text(self):
        documentation = []
        
        documentation.append("The VBA script contains the following components:")

        if self.comments:
            documentation.append("\n**Comments:**")
            for comment in self.comments:
                documentation.append(f"- {comment.strip()}")

        if self.variables:
            documentation.append("\n**Variables:**")
            for var in self.variables:
                name, var_type = var
                documentation.append(f"- Variable '{name}' of type '{var_type}'")

        if self.functions:
            documentation.append("\n**Functions:**")
            for func in self.functions:
                documentation.append(f"- Function '{func['name']}' with parameters '{func['params']}' returning '{func['return_type']}'")
                documentation.append(f"  Function Body: {func['body']}")

        if self.subroutines:
            documentation.append("\n**Subroutines:**")
            for sub in self.subroutines:
                documentation.append(f"- Subroutine '{sub['name']}' with parameters '{sub['params']}'")
                documentation.append(f"  Subroutine Body: {sub['body']}")

        if self.loops:
            documentation.append("\n**Loops:**")
            for loop in self.loops:
                documentation.append(f"- Loop: {loop.strip()}")

        if self.conditionals:
            documentation.append("\n**Conditionals:**")
            for cond in self.conditionals:
                documentation.append(f"- Conditional: {cond.strip()}")

        if self.error_handling:
            documentation.append("\n**Error Handling:**")
            for err in self.error_handling:
                documentation.append(f"- Error Handling: {err.strip()}")

        if self.classes:
            documentation.append("\n**Classes:**")
            for cls in self.classes:
                documentation.append(f"- Class: {cls.strip()}")

        if self.errors:
            documentation.append("\n**Errors Detected:**")
            for err in self.errors:
                documentation.append(f"- {err}")

        return "\n".join(documentation)

    def _generate_chatgpt_response(self, documentation, openai_api_key):
        openai.api_key = openai_api_key
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"Generate human-readable documentation for the following VBA script:\n{documentation}"}
            ]
        )
        return response['choices'][0]['message']['content']

# Test case for the enhanced VBA parser with ChatGPT integration

vba_code_with_errors = """
' This is a comment
Dim count As Integer
Dim unusedVar As String

Function AddNumbers(a As Integer, b As Integer) As Integer
    Dim result As Integer
    result = a + b
    AddNumbers = result
End Function

Sub ShowMessage(msg As String)
    MsgBox msg
End Sub

Sub LoopExample()
    Dim i As Integer
    For i = 1 To 10
        MsgBox i
Next i

Sub ConditionalExample(x As Integer)
    If x > 0 Then
        MsgBox "Positive"
    Else
        MsgBox "Non-positive"
    ' Missing End If here

Sub ErrorExample()
    On Error Resume Next
    Dim x As Integer
    x = 1 / 0
    If Err.Number <> 0 Then
        MsgBox "Error occurred"
    End If
End Sub

Class MyClass
    Private x As Integer
    Public Property Get Value() As Integer
        Value = x
    End Property
    Public Property Let Value(newValue As Integer)
        x = newValue
    End Property
End Class

' Use of undeclared variable
total = AddNumbers(5, 10)
"""

# Assuming the enhanced VBAParser class code is in vba_parser.py
parser = VBAParser(vba_code_with_errors)

# Parse the VBA code
parser.parse()

# Replace 'your_openai_api_key' with your actual OpenAI API key
openai_api_key = 'paste_your_api_key_to_proceed'

# Get the human-readable documentation
human_readable_docs = parser.get_human_readable_documentation(openai_api_key)

# Print the generated documentation
print(human_readable_docs)