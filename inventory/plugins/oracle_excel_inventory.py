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
    - Diseñado específicamente para el proyecto checklist_databases
    - Mantiene compatibilidad total con el código existente
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
    environment_filter:
        description: Filtrar por ambiente específico
        type: str
        required: false
    company_filter:
        description: Filtrar por compañía específica  
        type: str
        required: false
    datacenter_filter:
        description: Filtrar por datacenter específico
        type: str
        required: false
'''

EXAMPLES = '''
# Configuración básica
---
plugin: oracle_excel_inventory
file_path: "inventory/InventarioBD.xlsx"
sheet_name: "ORACLE"

# Con filtros desde AWX
---
plugin: oracle_excel_inventory
file_path: "inventory/InventarioBD.xlsx"
environment_filter: "{{ target_environment | default('') }}"
company_filter: "{{ target_company | default('') }}"
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
        
        # Extraer parámetros
        file_path = config.get('file_path')
        sheet_name = config.get('sheet_name', 'ORACLE')
        environment_filter = config.get('environment_filter', '').strip()
        company_filter = config.get('company_filter', '').strip()
        datacenter_filter = config.get('datacenter_filter', '').strip()

        # Validar archivo
        if not file_path:
            raise AnsibleParserError("file_path es requerido")
        
        if not os.path.exists(file_path):
            raise AnsibleParserError(f"Archivo Excel no encontrado: {file_path}")

        # Procesar Excel
        oracle_databases = self._process_excel(
            file_path, sheet_name, 
            environment_filter, company_filter, datacenter_filter
        )

        # Generar inventario
        self._create_inventory(oracle_databases)

    def _process_excel(self, file_path: str, sheet_name: str, 
                      env_filter: str, company_filter: str, 
                      datacenter_filter: str) -> Dict[str, Any]:
        """Procesar archivo Excel y generar oracle_databases"""
        
        try:
            # Leer Excel
            display.vvv(f"Leyendo Excel: {file_path}, hoja: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            
            # Verificar columnas requeridas
            required_columns = [
                'Instancia', 'Host', 'Direccion_IP', 'Cadena_TNS',
                'Ambiente', 'Motor_BD', 'Sistema_Operativo', 
                'Release_SO', 'Release_BD'
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise AnsibleParserError(
                    f"Columnas faltantes en Excel: {', '.join(missing_columns)}\n"
                    f"Columnas disponibles: {', '.join(df.columns.tolist())}"
                )

            # Aplicar filtros
            if env_filter:
                df = df[df['Ambiente'].str.strip() == env_filter]
                display.vvv(f"Filtro ambiente '{env_filter}': {len(df)} filas")

            if company_filter:
                df = df[df['Companía'].str.strip() == company_filter]
                display.vvv(f"Filtro compañía '{company_filter}': {len(df)} filas")
                
            if datacenter_filter:
                df = df[df['Datacenter'].str.strip() == datacenter_filter]
                display.vvv(f"Filtro datacenter '{datacenter_filter}': {len(df)} filas")

            # Construir oracle_databases
            oracle_databases = {}
            
            for _, row in df.iterrows():
                # Usar Cadena TNS como clave (más específico que instancia)
                cadena_tns = str(row['Cadena_TNS']).strip()
                
                if not cadena_tns or cadena_tns.lower() in ['nan', 'none', '']:
                    continue  # Saltar filas sin cadena TNS válida

                # Construir entrada de base de datos
                oracle_databases[cadena_tns] = {
                    # Datos de conexión requeridos
                    'instancia': str(row['Instancia']).strip(),
                    'host': str(row['Host']).strip(),
                    'direccion_ip': str(row['Direccion_IP']).strip(),
                    'cadena_tns': cadena_tns,
                    'puerto': 1521,  # Puerto por defecto Oracle
                    
                    # Metadatos del Excel
                    'motor_bd': str(row['Motor_BD']).strip(),
                    'sistema_operativo': str(row['Sistema_Operativo']).strip(),
                    'release_so': str(row['Release_SO']).strip(),
                    'release_bd': str(row['Release_BD']).strip(),
                    'ambiente': str(row['Ambiente']).strip(),
                    
                    # Metadatos adicionales si existen
                    'datacenter': str(row.get('Datacenter', 'Unknown')).strip(),
                    'compania': str(row.get('Companía', 'Unknown')).strip(),
                    'dominio': str(row.get('Dominio', 'Unknown')).strip(),
                    'transversal': str(row.get('Transversal', 'No')).strip()
                }

            display.vvv(f"Oracle databases generadas: {len(oracle_databases)}")
            display.vvv(f"Instancias: {list(oracle_databases.keys())}")
            
            return oracle_databases

        except Exception as e:
            raise AnsibleError(f"Error procesando Excel: {str(e)}")

    def _create_inventory(self, oracle_databases: Dict[str, Any]) -> None:
        """Crear inventario Ansible con oracle_databases"""
        
        # Crear host especial para variables globales
        host_name = 'localhost'
        self.inventory.add_host(host_name)
        
        # Asignar oracle_databases como variable del host
        self.inventory.set_variable(host_name, 'oracle_databases', oracle_databases)
        
        # Crear grupos por ambiente
        ambientes = {}
        for key, db in oracle_databases.items():
            ambiente = db['ambiente'].lower().replace(' ', '_').replace('ó', 'o')
            
            if ambiente not in ambientes:
                ambientes[ambiente] = []
            ambientes[ambiente].append(key)

        # Crear grupos de ambiente en inventario
        for ambiente, instancias in ambientes.items():
            group_name = f"env_{ambiente}"
            self.inventory.add_group(group_name)
            
            # Agregar host a cada grupo (para poder usar groups en playbooks)
            self.inventory.add_child(group_name, host_name)
            
            # También crear variable con las instancias del ambiente
            self.inventory.set_variable(group_name, 'instances', instancias)

        # Crear grupos adicionales útiles
        companies = {}
        datacenters = {}
        
        for key, db in oracle_databases.items():
            # Grupos por compañía
            company = db['compania'].lower().replace(' ', '_')
            if company not in companies:
                companies[company] = []
            companies[company].append(key)
            
            # Grupos por datacenter  
            datacenter = db['datacenter'].lower().replace(' ', '_')
            if datacenter not in datacenters:
                datacenters[datacenter] = []
            datacenters[datacenter].append(key)

        # Crear grupos de compañía
        for company, instancias in companies.items():
            if company != 'unknown':
                group_name = f"company_{company}"
                self.inventory.add_group(group_name)
                self.inventory.add_child(group_name, host_name)
                self.inventory.set_variable(group_name, 'instances', instancias)

        # Crear grupos de datacenter
        for datacenter, instancias in datacenters.items():
            if datacenter != 'unknown':
                group_name = f"dc_{datacenter}"
                self.inventory.add_group(group_name)
                self.inventory.add_child(group_name, host_name)
                self.inventory.set_variable(group_name, 'instances', instancias)

        display.vvv(f"Inventario creado con {len(oracle_databases)} instancias Oracle")
        display.vvv(f"Grupos creados: {len(ambientes)} ambientes, {len(companies)} compañías, {len(datacenters)} datacenters")