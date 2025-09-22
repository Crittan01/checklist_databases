#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Plugin de Inventario Oracle específico para checklist_databases
Genera la estructura oracle_databases desde Excel
"""

from __future__ import (absolute_import, division, print_function)
__metaclass__ = type

DOCUMENTATION = '''
---
name: oracle_excel_inventory
author: DevOps Team
version_added: "1.0.0"
short_description: Plugin específico para inventario Oracle desde Excel
description:
    - Genera la estructura oracle_databases desde archivos Excel
    - Carga todos los datos como variables de inventario
    - Los filtros se aplican dinámicamente en playbooks
requirements:
    - pandas >= 1.0.0
    - openpyxl >= 3.0.0
options:
    plugin:
        description: Token del plugin
        type: str
        required: true
        choices: ['oracle_excel_inventory']
    file_path:
        description: Ruta al archivo Excel
        type: path
        required: true
    sheet_name:
        description: Nombre de la hoja Excel (opcional)
        type: str
        required: false
        default: "ORACLE"
'''

EXAMPLES = '''
# Configuración básica - Carga todo el Excel
---
plugin: oracle_excel_inventory
file_path: "inventory/InventarioBD.xlsx"
sheet_name: "ORACLE"

# Los filtros se aplican dinámicamente en playbooks usando survey variables
'''

import os
import sys
from typing import Dict, Any, Optional

from ansible.plugins.inventory import BaseInventoryPlugin, Constructable
from ansible.errors import AnsibleError, AnsibleParserError
from ansible.utils.display import Display

display = Display()

# Verificar pandas
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


class InventoryModule(BaseInventoryPlugin, Constructable):
    """Plugin de inventario Oracle específico"""

    NAME = 'oracle_excel_inventory'

    def verify_file(self, path: str) -> bool:
        """Verificar que el archivo sea válido para este plugin"""
        if super().verify_file(path):
            if path.endswith(('oracle_excel_inventory.yaml', 'oracle_excel_inventory.yml')):
                return True
        return False

    def parse(self, inventory, loader, path: str, cache: bool = True) -> None:
        """Parsear el inventario y generar oracle_databases"""
        super().parse(inventory, loader, path, cache)

        # Verificar dependencias
        if not HAS_PANDAS:
            raise AnsibleError(
                "Se requiere pandas para procesar Excel. "
                "Instalar con: pip install pandas openpyxl"
            )

        # Leer configuración
        config = self._read_config_data(path)
        
        # Extraer parámetros básicos (sin filtros)
        file_path = config.get('file_path')
        sheet_name = config.get('sheet_name', 'ORACLE')

        # Validar archivo
        if not file_path:
            raise AnsibleParserError("file_path es requerido")
        
        if not os.path.exists(file_path):
            raise AnsibleParserError(f"Archivo Excel no encontrado: {file_path}")

        # Procesar Excel (sin filtros - carga todo)
        oracle_data = self._process_excel(file_path, sheet_name)

        # Generar inventario con todos los datos
        self._create_inventory(oracle_data)

    def _process_excel(self, file_path: str, sheet_name: str) -> Dict[str, Any]:
        """Procesar archivo Excel y generar oracle_databases completo"""
        
        try:
            # Leer Excel
            display.vvv(f"Leyendo Excel: {file_path}, hoja: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            
            # Verificar columnas requeridas
            required_columns = [
                'Instancia', 'Host', 'Dirección IP', 'Cadena TNS',
                'Ambiente', 'Motor BD', 'Sistema Operativo', 
                'Release SO', 'Release BD'
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise AnsibleParserError(
                    f"Columnas faltantes en Excel: {', '.join(missing_columns)}\n"
                    f"Columnas disponibles: {', '.join(df.columns.tolist())}"
                )

            # NO aplicar filtros - cargar todo
            display.vvv(f"Procesando {len(df)} filas del Excel")

            # Construir oracle_databases completo
            oracle_databases = {}
            
            # Metadatos para filtros dinámicos
            environments = set()
            companies = set()
            datacenters = set()
            hosts = set()
            
            for _, row in df.iterrows():
                # Usar Cadena TNS como clave
                cadena_tns = str(row['Cadena TNS']).strip()
                
                if not cadena_tns or cadena_tns.lower() in ['nan', 'none', '']:
                    continue

                # Construir entrada de base de datos
                oracle_databases[cadena_tns] = {
                    # Datos de conexión requeridos
                    'instancia': str(row['Instancia']).strip(),
                    'host': str(row['Host']).strip(),
                    'direccion_ip': str(row['Dirección IP']).strip(),
                    'cadena_tns': cadena_tns,
                    'puerto': 1521,
                    'usuario': 'ansible_user',
                    'password': 'ansibledb123*',
                    
                    # Metadatos del Excel
                    'motor_bd': str(row['Motor BD']).strip(),
                    'sistema_operativo': str(row['Sistema Operativo']).strip(),
                    'release_so': str(row['Release SO']).strip(),
                    'release_bd': str(row['Release BD']).strip(),
                    'ambiente': str(row['Ambiente']).strip(),
                    'datacenter': str(row.get('Datacenter', 'Unknown')).strip(),
                    'compania': str(row.get('Compañía', 'Unknown')).strip(),
                    'dominio': str(row.get('Dominio', 'Unknown')).strip(),
                    'transversal': str(row.get('Transversal', 'No')).strip()
                }
                
                # Recopilar metadatos para filtros
                environments.add(str(row['Ambiente']).strip())
                companies.add(str(row.get('Compañía', 'Unknown')).strip())
                datacenters.add(str(row.get('Datacenter', 'Unknown')).strip())
                hosts.add(str(row['Host']).strip())

            # Preparar datos completos
            oracle_data = {
                'oracle_databases': oracle_databases,
                'metadata': {
                    'total_instances': len(oracle_databases),
                    'environments': sorted(list(environments)),
                    'companies': sorted(list(companies)),
                    'datacenters': sorted(list(datacenters)),
                    'hosts': sorted(list(hosts)),
                    'source_file': file_path,
                    'last_loaded': 'TIMESTAMP_PLACEHOLDER'
                }
            }

            display.vvv(f"Oracle databases cargadas: {len(oracle_databases)}")
            display.vvv(f"Ambientes: {sorted(list(environments))}")
            display.vvv(f"Compañías: {sorted(list(companies))}")
            
            return oracle_data

        except Exception as e:
            raise AnsibleError(f"Error procesando Excel: {str(e)}")

    def _create_inventory(self, oracle_data: Dict[str, Any]) -> None:
        """Crear inventario Ansible con oracle_databases y metadatos"""
        
        oracle_databases = oracle_data['oracle_databases']
        metadata = oracle_data['metadata']
        
        # Agregar timestamp real
        import datetime
        metadata['last_loaded'] = datetime.datetime.now().isoformat()
        
        # Crear host especial para variables globales
        host_name = 'localhost'
        self.inventory.add_host(host_name)
        
        # Asignar oracle_databases como variable del host (compatible con código existente)
        self.inventory.set_variable(host_name, 'oracle_databases', oracle_databases)
        
        # Asignar metadatos útiles para filtros dinámicos
        self.inventory.set_variable(host_name, 'oracle_metadata', metadata)
        self.inventory.set_variable(host_name, 'available_environments', metadata['environments'])
        self.inventory.set_variable(host_name, 'available_companies', metadata['companies'])
        self.inventory.set_variable(host_name, 'available_datacenters', metadata['datacenters'])
        self.inventory.set_variable(host_name, 'available_hosts', metadata['hosts'])

        # Crear grupos por ambiente (para organización en AWX)
        for ambiente in metadata['environments']:
            ambiente_clean = ambiente.lower().replace(' ', '_').replace('ó', 'o')
            group_name = f"env_{ambiente_clean}"
            self.inventory.add_group(group_name)
            self.inventory.add_child(group_name, host_name)

        # Crear grupos por compañía
        for company in metadata['companies']:
            if company != 'Unknown':
                company_clean = company.lower().replace(' ', '_')
                group_name = f"company_{company_clean}"
                self.inventory.add_group(group_name)
                self.inventory.add_child(group_name, host_name)

        # Crear grupos por datacenter
        for datacenter in metadata['datacenters']:
            if datacenter != 'Unknown':
                datacenter_clean = datacenter.lower().replace(' ', '_')
                group_name = f"dc_{datacenter_clean}"
                self.inventory.add_group(group_name)
                self.inventory.add_child(group_name, host_name)

        display.vvv(f"Inventario creado:")
        display.vvv(f"  - {len(oracle_databases)} instancias Oracle")
        display.vvv(f"  - {len(metadata['environments'])} ambientes")
        display.vvv(f"  - {len(metadata['companies'])} compañías")
        display.vvv(f"  - Variables disponibles: oracle_databases, oracle_metadata")
