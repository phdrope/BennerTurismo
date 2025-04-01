import xml.etree.ElementTree as ET
import re

def find_nrovoo_tags(xml_string):
    """Encontra todas as tags <NroVoo>...</NroVoo> no XML e retorna a tag completa se o valor for diferente de 4 dígitos numéricos."""
    matches = re.findall(r'(<NroVoo>.*?</NroVoo>)', xml_string, re.DOTALL)
    
    # Filtra tags que não tenham exatamente 4 dígitos numéricos dentro
    filtered_tags = [tag for tag in matches if not re.fullmatch(r'<NroVoo>\s*\d{4}\s*</NroVoo>', tag.strip())]
    
    return filtered_tags

# Exemplo de uso
xml_string = (SELECT XMLRESERVA FROM BB_LOGINTEGRACOES WHERE HANDLE = :HANDLE)

nrovoo_tags = find_nrovoo_tags(xml_string)
print(nrovoo_tags)
