import json
import openpyxl

def parse_openapi_to_excel(openapi_path, excel_path):
    def resolve_ref(ref, openapi_data):
        """Resolve a $ref path in the OpenAPI document."""
        path = ref.lstrip("#/").split("/")
        schema = openapi_data
        for part in path:
            schema = schema.get(part, {})
        return schema

    def process_properties(properties, required_fields, parent="", openapi_data=None):
        rows = []
        for field_name, field_data in properties.items():
            full_name = f"{parent}.{field_name}" if parent else field_name

            # Resolver $ref se existir
            if "$ref" in field_data:
                field_data = resolve_ref(field_data["$ref"], openapi_data)

            field_type = field_data.get("type", "unknown")
            description = field_data.get("description", "No description")
            required = is_field_required(required_fields, field_name)    

            # Adicionar campo atual
            rows.append([full_name, field_type, required, description])
            
            # Processar objetos aninhados
            if field_type == "object" and "properties" in field_data:
                nested_properties = field_data["properties"]
                nested_required = field_data.get("required", [])
                rows.extend(process_properties(nested_properties, nested_required, parent=full_name, openapi_data=openapi_data))
            
            # Processar arrays
            if field_type == "array" and "items" in field_data:
                items = field_data["items"]
                if "$ref" in items:
                    items = resolve_ref(items["$ref"], openapi_data)
                if "properties" in items:
                    nested_properties = items["properties"]
                    nested_required = items.get("required", [])
                    rows.extend(process_properties(nested_properties, nested_required, parent=f"{full_name}[]", openapi_data=openapi_data))
        return rows

    def is_field_required(required_fields, field_name):
        if field_name in required_fields: 
            required = "true"
        else:
            required = "false"
        return required

    # Carregar o arquivo OpenAPI
    with open(openapi_path, "r", encoding="utf-8") as file:
        openapi_data = json.load(file)
    
    # Criar um workbook do Excel
    wb = openpyxl.Workbook()

    # Remover a aba inicial padrão
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Percorrer os endpoints
    for endpoint, methods in openapi_data.get("paths", {}).items():
        for method, details in methods.items():
            # Ignorar métodos sem requestBody
            request_body = details.get("requestBody", {}).get("content", {})
            if not request_body:
                continue

            # Obter schema de propriedades
            for content_type, content_data in request_body.items():
                schema = content_data.get("schema", {})
                
                # Resolver $ref no schema principal, se necessário
                if "$ref" in schema:
                    schema = resolve_ref(schema["$ref"], openapi_data)
                
                properties = schema.get("properties", {})
                required_fields = schema.get("required", [])

                # Criar uma nova aba para o serviço
                summary = details.get("summary", "No summary")[:30]
                tab_name = f"{summary}"
                tab_name = tab_name[:31]  # Limitar a 31 caracteres (limite do Excel)
                ws = wb.create_sheet(title=tab_name)

                # Adicionar a tabela de informações do serviço
                service_info = [
                    ["Endpoint", "Method", "Summary"],
                    [endpoint, method.upper(), summary]
                ]
                for row in service_info:
                    ws.append(row)

                # Adicionar uma linha em branco para separar a tabela de serviço da tabela de campos
                ws.append([])

                # Adicionar cabeçalhos para a tabela de campos
                headers = ["Field Name", "Type", "Required", "Description"]
                ws.append(headers)

                # Processar propriedades e adicionar à aba
                rows = process_properties(properties, required_fields, openapi_data=openapi_data)
                for row in rows:
                    ws.append(row)

    # Salvar o Excel
    wb.save(excel_path)
    print(f"Arquivo Excel salvo em: {excel_path}")

# Caminhos para o arquivo OpenAPI e o Excel de saída
openapi_path = "openapi.json"
excel_path = "openapi_as_exel.xlsx"

# Gerar o Excel
parse_openapi_to_excel(openapi_path, excel_path)
