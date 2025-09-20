
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright: (c) 2025, RandyRozo
# GNU General Public License v3.0+

"""
Excel Inventory Plugin for Ansible

This inventory plugin generates dynamic Ansible inventories from Excel files,
supporting both local files and SharePoint Online integration.

Author: Randy Rozo (@randyrozo)
License: GPL v3.0+
"""

from __future__ import (absolute_import, division, print_function)
__metaclass__ = type


DOCUMENTATION = '''
---
name: randyrozo.utils.excel_inventory
author:
    - Randy Rozo (@randyrozo)
version_added: "1.0.0"
short_description: Dynamic inventory plugin for Excel files (local and SharePoint)
description:
    - Generates dynamic Ansible inventories from Excel files stored locally or in SharePoint Online
    - Supports advanced filtering with multiple criteria and Jinja2 templates
    - Provides flexible automatic and conditional grouping capabilities
    - Full integration with AWX/AAP for dynamic inventory management
    - Customizable host variable mapping from Excel columns
    - Professional error handling with clear diagnostic messages
requirements:
    - python >= 3.6
    - pandas >= 1.0.0
    - openpyxl >= 3.0.0
    - Office365-REST-Python-Client >= 2.3.0 (only for SharePoint functionality)
extends_documentation_fragment:
    - constructed
    - inventory_cache
notes:
    - Configuration files must end with C(.excel_inventory.yml) or C(.excel_inventory.yaml)
    - If sheet_name is not specified, the first sheet in the Excel file will be used
    - Group names are automatically normalized (lowercase, spaces converted to underscores)
    - For SharePoint access, you need a registered Azure AD application with appropriate permissions
    - All Jinja2 templates have access to variables passed from AWX/AAP surveys
    - The plugin supports both .xlsx and .xls file formats
    - Empty hostname rows are automatically skipped unless skip_empty_hosts is disabled
seealso:
    - name: ansible.builtin.constructed inventory plugin
      description: The constructed inventory plugin reference
      link: https://docs.ansible.com/ansible/latest/collections/ansible/builtin/constructed_inventory.html
    - name: Ansible inventory plugins
      description: Generic guide for inventory plugins
      link: https://docs.ansible.com/ansible/latest/plugins/inventory.html

options:
    plugin:
        description: 
            - Token that ensures this is a source file for the excel_inventory plugin
        type: str
        required: true
        choices: ['excel_inventory', 'randyrozo.utils.excel_inventory']

    source_type:
        description: 
            - Type of data source for the Excel file
            - Use C(local) for files stored on the local filesystem
            - Use C(sharepoint) for files stored in SharePoint Online
        type: str
        required: true
        choices: ['local', 'sharepoint']
        
    local_config:
        description: 
            - Configuration options for local Excel files
            - Required when I(source_type=local)
        type: dict
        required: false
        suboptions:
            file_path:
                description: 
                    - Full absolute or relative path to the Excel file
                    - Supports both .xlsx and .xls formats
                    - Can include Jinja2 templates for dynamic paths
                type: path
                required: true

    sharepoint_config:
        description: 
            - Configuration options for SharePoint Online files
            - Required when I(source_type=sharepoint)
            - Requires a registered Azure AD application
        type: dict
        required: false
        suboptions:
            site_url:
                description: 
                    - Complete URL of the SharePoint site
                    - Example C(https://company.sharepoint.com/sites/infrastructure)
                type: str
                required: true
            folder_path:
                description: 
                    - Folder path within the SharePoint site (relative to site)
                    - Example C(/sites/infrastructure/Shared Documents/Inventory)
                type: str
                required: true
            filename:
                description: 
                    - Name of the Excel file in SharePoint
                    - Must include file extension (.xlsx or .xls)
                type: str
                required: true
            client_id:
                description: 
                    - Azure AD application client ID (also called application ID)
                    - Can be obtained from Azure Portal > App registrations
                type: str
                required: true
            client_secret:
                description: 
                    - Azure AD application client secret
                    - Should be stored securely and passed via encrypted variables
                type: str
                required: true
                no_log: true
            tenant_id:
                description: 
                    - Azure AD tenant ID (also called directory ID)
                    - Can be found in Azure Portal > Azure Active Directory > Properties
                type: str
                required: true

    sheet_name:
        description: 
            - Name of the Excel sheet/worksheet to process
            - If not specified, the first sheet will be used
            - Case-sensitive sheet name
        type: str
        required: false

    hostname_column:
        description: 
            - Name of the Excel column containing hostnames/server names
            - This column will be used as the inventory hostname
            - Case-sensitive column name
        type: str
        default: hostname
        required: false

    filters:
        description: 
            - Dictionary of filters to include only specific hosts
            - All filters are applied with AND logic (all must match)
            - Supports Jinja2 templates for dynamic filtering from AWX surveys
            - Use empty string or omit filter to disable filtering for that column
        type: dict
        default: {}
        required: false
        
    compose:
        description: 
            - Dictionary mapping host variable names to Excel column names
            - Key is the desired Ansible variable name, value is the Excel column name
            - Only variables specified here will be created for each host
            - Provides complete control over which Excel data becomes host variables
        type: dict
        default: {}
        required: false
        
    keyed_groups:
        description: 
            - List of automatic group configurations based on unique column values
            - Creates one inventory group for each unique value found in the specified column
            - Supports both simple column name format and advanced configuration with prefixes
        type: list
        elements: dict
        default: []
        required: false
        suboptions:
            key:
                description: 
                    - Name of the Excel column to group by
                    - Must be an existing column in the Excel file
                type: str
                required: true
            prefix:
                description: 
                    - Optional prefix for group names
                    - Helps avoid naming conflicts between different grouping categories
                    - Example C(env) creates groups like C(env_production), C(env_development)
                type: str
                required: false
                default: ""
            separator:
                description: 
                    - Character(s) used between prefix and base group name
                    - Only used when prefix is specified
                type: str
                required: false
                default: "_"

    groups:
        description: 
            - Dictionary of conditional groups using Python-like expressions
            - Key is the group name, value is a safe Python expression
            - Expressions support comparison operators (==, !=, <, <=, >, >=)
            - Expressions support logical operators (and, or, not)
            - Expressions support membership operators (in, not in)
            - Variables in expressions must match Excel column names exactly
        type: dict
        default: {}
        required: false

    skip_empty_hosts:
        description: 
            - Whether to skip Excel rows where the hostname column is empty or null
            - When enabled, rows with empty hostnames are silently ignored
            - When disabled, may cause errors if hostname column contains empty values
        type: bool
        default: true
        required: false

    strict:
        description: 
            - Enable strict mode for Jinja2 template processing
            - When enabled, template errors cause the plugin to fail
            - When disabled, template errors generate warnings and use default values
        type: bool
        default: false
        required: false
'''

EXAMPLES = '''
# Minimal local file example
---
plugin: randyrozo.utils.excel_inventory
source_type: local
local_config:
  file_path: "./servers.xlsx"
hostname_column: "ServerName"
compose:
  ansible_host: "IP_Address"

# Complete local file example with filtering and grouping
---
plugin: randyrozo.utils.excel_inventory
source_type: local
local_config:
  file_path: "/inventory/production_servers.xlsx"
hostname_column: "HostName"
sheet_name: "Production"

# Dynamic filtering using AWX survey variables
filters:
  Environment: "{{ target_env | default('') }}"
  Project: "{{ project_code | default('') }}"
  Status: "Active"

# Comprehensive variable mapping
compose:
  ansible_host: "Management_IP"
  ansible_port: "SSH_Port"
  ansible_user: "Admin_User"
  server_role: "Server_Role"
  environment: "Environment"
  project_code: "Project"
  cost_center: "Cost_Center"
  datacenter: "DC_Location"

# Automatic grouping with prefixes
keyed_groups:
  - key: Environment
    prefix: env
  - key: Server_Role
    prefix: role
  - key: DC_Location
    prefix: datacenter
    separator: "-"

# Conditional grouping with complex expressions
groups:
  production: "Environment == 'Production'"
  development: "Environment in ['Development', 'Dev', 'Test']"
  web_servers: "Server_Role == 'Web Server'"
  database_servers: "Server_Role == 'Database'"
  high_availability: "Environment == 'Production' and Server_Role in ['Database', 'Load Balancer']"
  large_instances: "CPU_Count >= 8 and Memory_GB >= 32"

# SharePoint Online integration example
---
plugin: randyrozo.utils.excel_inventory
source_type: sharepoint
sharepoint_config:
  site_url: "https://company.sharepoint.com/sites/infrastructure"
  folder_path: "/sites/infrastructure/Shared Documents/Inventories"
  filename: "master_inventory.xlsx"
  client_id: "{{ vault_sharepoint_client_id }}"
  client_secret: "{{ vault_sharepoint_client_secret }}"
  tenant_id: "{{ vault_sharepoint_tenant_id }}"

hostname_column: "FQDN"
sheet_name: "Current_Infrastructure"
skip_empty_hosts: true
strict: false

compose:
  ansible_host: "Primary_IP"
  backup_ip: "Secondary_IP"
  os_family: "Operating_System"
  patch_group: "Patching_Schedule"

# AWX/AAP survey integration example
---
plugin: randyrozo.utils.excel_inventory
source_type: local
local_config:
  file_path: "{{ survey_inventory_file }}"

hostname_column: "{{ survey_hostname_column | default('HostName') }}"

# Dynamic filters from survey
filters:
  Environment: "{{ survey_environment | default('') }}"
  Application: "{{ survey_application | default('') }}"
  Maintenance_Window: "{{ survey_maintenance_window | default('') }}"

compose:
  ansible_host: "{{ survey_ip_column | default('IP_Address') }}"
  environment: "Environment"
  application: "Application"

keyed_groups:
  - key: "{{ survey_group_by_column | default('Environment') }}"

# Multi-environment example
---
plugin: randyrozo.utils.excel_inventory
source_type: local
local_config:
  file_path: "./multi_env_inventory.xlsx"
hostname_column: "Host"

filters:
  Status: "Online"
  
compose:
  ansible_host: "IP"
  env: "Environment"
  tier: "Application_Tier"

keyed_groups:
  - key: Environment
  - key: Application_Tier
    prefix: tier
  - key: Operating_System
    prefix: os

groups:
  # Environment-based groups
  prod: "Environment == 'Production'"
  nonprod: "Environment != 'Production'"
  
  # Architecture-based groups
  web_tier: "Application_Tier == 'Web'"
  app_tier: "Application_Tier == 'Application'"
  db_tier: "Application_Tier == 'Database'"
  
  # OS-based groups
  linux: "Operating_System.startswith('Linux')"
  windows: "Operating_System.startswith('Windows')"
  
  # Complex conditional groups
  critical_prod: "Environment == 'Production' and Application_Tier in ['Database', 'Application']"
  patching_group_a: "Patch_Group == 'A' and Environment == 'Production'"
'''


import os
import sys
import ast
import tempfile
import operator
from typing import Dict, List, Any, Optional, Union

from ansible.plugins.inventory import BaseInventoryPlugin, Constructable, Cacheable
from ansible.errors import AnsibleError, AnsibleParserError
from ansible.config.manager import ensure_type
from ansible.utils.display import Display
from ansible.template import Templar

display = Display()

# Verificar dependencias
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from office365.runtime.auth.client_credential import ClientCredential
    from office365.sharepoint.client_context import ClientContext
    HAS_OFFICE365 = True
except ImportError:
    HAS_OFFICE365 = False


class InventoryModule(BaseInventoryPlugin, Constructable, Cacheable):
    """
    Excel Inventory Plugin for Ansible.
    
    Generates dynamic inventories from Excel files stored locally or in SharePoint Online.
    Supports advanced filtering, flexible grouping, and AWX/AAP integration.
    
    Attributes:
        NAME (str): Plugin identifier for Ansible
        template_handle (Templar): Jinja2 template processor
        source_type (str): Type of data source ('local' or 'sharepoint')
        sheet_name (Optional[str]): Excel sheet name to process
        local_config (Dict[str, Any]): Configuration for local files
        sharepoint_config (Dict[str, Any]): Configuration for SharePoint files
        hostname_column (str): Column name containing hostnames
        filters (Dict[str, Any]): Filters to apply to hosts
        compose (Dict[str, str]): Variable mapping configuration
        keyed_groups (List[Dict[str, Any]]): Automatic grouping configuration
        groups (Dict[str, str]): Conditional grouping configuration
        skip_empty_hosts (bool): Whether to skip empty hostname rows
        strict (bool): Strict mode for template processing
    """

    NAME = 'randyrozo.utils.excel_inventory'
    # NAME = 'excel_inventory'

    def __init__(self) -> None:
        """
        Initialize the Excel inventory plugin.
        
        Raises:
            AnsibleError: If Python version is less than 3.6
        """
        if sys.version_info < (3, 6):
            py_ver = sys.version_info[0]
            raise AnsibleError(
                f"Unsupported Python version {py_ver}, "
                f"supported python version is 3.6 and above"
            )

        super().__init__()


    def verify_file(self, path: str) -> bool:
        """
        Verify that the given path is valid for this plugin.
        
        Args:
            path (str): Path to the inventory configuration file
            
        Returns:
            bool: True if the file is valid for this plugin
            
        Raises:
            AnsibleParserError: If file extension is not .excel_inventory.yml/yaml
        """
        
        if super().verify_file(path):
            if path.endswith(('excel_inventory.yaml', 'excel_inventory.yml')):
                return True
            raise AnsibleParserError(
                "Path is not valid. All excel inventory sources must have "
                "a suffix of .excel_inventory.yml or .excel_inventory.yaml."
            )
        return False


    def parse(self, inventory, loader, path: str, cache: bool = True) -> None:
        """
        Parse the inventory source and populate the inventory.
        
        This is the main entry point for the plugin. It reads configuration,
        processes the Excel file, and populates the Ansible inventory with
        hosts, variables, and groups.
        
        Args:
            inventory: Ansible inventory object to populate
            loader: Ansible data loader
            path (str): Path to the inventory configuration file
            cache (bool, optional): Whether to use caching. Defaults to True.
            
        Raises:
            AnsibleError: If requirements are not met or processing fails
        """
        super().parse(inventory, loader, path, cache)

        # Initialize Jinja2 template processor
        self.template_handle = Templar(loader=loader)

        # Verify all dependencies are available
        self._check_requirements()

        # Read and validate configuration
        self._configure(path)

        # Process Excel file based on source type
        if self.source_type == 'local':
            self._process_local_source()
        elif self.source_type == 'sharepoint':
            self._process_sharepoint_source()


    def _check_requirements(self) -> None:
        """
        Verify that all required dependencies are installed.
        
        Raises:
            AnsibleError: If pandas is not available
        """
        if not HAS_PANDAS:
            raise AnsibleError(
                "Se requiere 'pandas' para procesar archivos Excel. "
                "Instala con: pip install pandas openpyxl"
            )

    def _configure(self, path: str) -> None:
        """
        Read and validate plugin configuration.
        
        Reads the YAML configuration file, renders any Jinja2 templates,
        validates all settings, and sets instance attributes.
        
        Args:
            path (str): Path to the configuration file
            
        Raises:
            AnsibleParserError: If configuration is invalid
            AnsibleError: If template rendering fails
        """
        # Read raw configuration
        config = self._read_config_data(path)

        # Render Jinja2 templates in configuration
        config = self._render_templates(config)
        
        # Set configuration attributes with validation
        arg = {
            'source_type': {
                'type': 'str', 
                'choices': ['local', 'sharepoint'], 
                'value': config.get('source_type', ''), 
                'required': True
            },
            'sheet_name': {
                'type': 'str', 
                'value': config.get('sheet_name', ''), 
                'required': False
            },
            'local_config': {
                'type': 'dict', 
                'value': config.get("local_config", {}), 
                'required': False
            },
            'sharepoint_config': {
                'type': 'dict', 
                'value': config.get("sharepoint_config", {}), 
                'required': False
            },
            'hostname_column': {
                'type': 'str', 
                'value': config.get('hostname_column', 'hostname'), 
                'required': False
            },
            'filters': {
                'type': 'dict', 
                'value': config.get("filters", {})
            },
            'compose': {
                'type': 'dict', 
                'value': config.get("compose", {})
            },
            'keyed_groups': {
                'type': 'list', 
                'value': config.get("keyed_groups", [])
            },
            'groups': {
                'type': 'dict', 
                'value': config.get("groups", {})
            },
            'skip_empty_hosts': {
                'type': 'bool', 
                'value': config.get("skip_empty_hosts", True), 
                'required': False
            },
            'strict': {
                'type': 'bool', 
                'value': config.get('strict', False), 
                'required': False
            }
        }

        # Validate and set all configuration attributes
        self.validate_and_set_args(arg)

        # Perform domain-specific validations
        self._validate_domain_specific()


    def _render_templates(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Recursively render Jinja2 templates in configuration values.
        
        Processes all string values in the configuration dictionary,
        rendering any Jinja2 templates with access to extra variables
        from AWX/APP surveys.
        
        Args:
            config (Dict[str, Any]): Raw configuration dictionary
            
        Returns:
            Dict[str, Any]: Configuration with rendered templates
            
        Raises:
            AnsibleError: If template rendering fails
        """
        def render_recursive(obj: Any) -> Any:
            """Recursively render templates in nested structures."""
            if isinstance(obj, dict):
                return {k: render_recursive(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [render_recursive(item) for item in obj]
            elif isinstance(obj, str):
                # Only render strings that contain template syntax (performance optimization)
                if '{{' in obj and '}}' in obj:
                    try:
                        return self.template_handle.template(obj)
                    except Exception as e:
                        error_msg = f"Template rendering failed for '{obj}': {str(e)}"
                        if self.strict:
                            raise AnsibleError(error_msg)
                        else:
                            display.warning(error_msg)
                            return obj
                return obj
            else:
                return obj

        try:
            return render_recursive(config)
        except Exception as e:
            raise AnsibleError(f"Error rendering configuration templates: {str(e)}")


    def validate_and_set_args(self, args: Dict[str, Dict[str, Any]]) -> None:
        """
        Validate configuration arguments and set instance attributes.
        
        Performs type validation, required field checking, and choice validation
        for all configuration parameters. Uses Ansible's ensure_type function
        for consistent type coercion and validation.
        
        Args:
            args (Dict[str, Dict[str, Any]]): Argument specifications containing:
                - type: Expected data type ('str', 'dict', 'list', 'bool')
                - value: Actual value from configuration
                - required: Whether the field is mandatory
                - choices: Valid values (for choice validation)
                
        Raises:
            AnsibleParserError: If validation fails for any parameter
        """

        for arg_name, arg_config in args.items():
            value = arg_config.get("value")
            arg_type = arg_config.get("type")
            is_required = arg_config.get("required", False)
            default_value = arg_config.get("default")

            # Check for required fields
            if is_required and value is None:
                raise AnsibleParserError(
                    f"Required configuration parameter '{arg_name}' is missing. "
                    f"Please check your inventory configuration file."
                )
            
            # Validate and set values
            if value is not None:
                try:
                    # Use Ansible's type validation
                    validated_value = ensure_type(value, arg_type)
                    
                    # Additional validation for string types
                    if arg_type == 'str':
                        validated_value = validated_value.strip()
                        
                        # Check for empty required strings
                        if arg_name in ['source_type', 'hostname_column'] and not validated_value:
                            raise ValueError(f"'{arg_name}' cannot be empty")
                        
                        # Validate choices if specified
                        if "choices" in arg_config:
                            choices = arg_config.get("choices")
                            if validated_value not in choices:
                                raise ValueError(
                                    f"Invalid value '{validated_value}' for '{arg_name}'. "
                                    f"Valid choices are: {', '.join(choices)}"
                                )

                    # Set the validated value as instance attribute
                    setattr(self, arg_name, validated_value)
                    
                except Exception as e:
                    raise AnsibleParserError(
                        f"Configuration validation failed for '{arg_name}': {str(e)}"
                    )


    def _validate_domain_specific(self) -> None:
        """
        Perform domain-specific configuration validations.
        
        Validates configuration requirements that are specific to Excel processing
        and SharePoint integration. These validations go beyond basic type checking
        to ensure the configuration makes sense for the intended use case.
        
        Local source validations:
        - Ensures local_config is provided and contains file_path
        - Validates file existence and accessibility
        - Checks file format (.xlsx or .xls)
        
        SharePoint source validations:
        - Ensures sharepoint_config contains all required fields
        - Validates Office365 library availability
        - Checks for required Azure AD application credentials
        
        Raises:
            AnsibleParserError: If domain-specific validation fails
            AnsibleError: If SharePoint library is missing when needed
        """
        
        if self.source_type == 'local':
            # Validate local file configuration
            if not self.local_config or not self.local_config.get('file_path'):
                raise AnsibleParserError(
                    "Missing required configuration for local source type.\n"
                    "Required configuration:\n"
                    "local_config:\n"
                    "  file_path: '/path/to/your/excel/file.xlsx'"
                )
            
            # Validate file existence and format
            file_path = self.local_config.get('file_path')
            if not os.path.exists(file_path):
                raise AnsibleParserError(
                    f"Excel file not found: {file_path}\n"
                    f"Please ensure the file exists and is accessible."
                )
            
            if not file_path.lower().endswith(('.xlsx', '.xls')):
                raise AnsibleParserError(
                    f"Unsupported file format: {file_path}\n"
                    f"Supported formats: .xlsx (Excel 2007+), .xls (Excel 97-2003)"
                )

        elif self.source_type == 'sharepoint':
            # Validate SharePoint library availability
            if not HAS_OFFICE365:
                raise AnsibleError(
                    "Missing required dependency for SharePoint integration.\n"
                    "Install with: pip install Office365-REST-Python-Client"
                )

            # Validate SharePoint configuration completeness
            if not self.sharepoint_config:
                raise AnsibleParserError(
                    "Missing required configuration for SharePoint source type.\n"
                    "Required configuration:\n"
                    "sharepoint_config:\n"
                    "  site_url: 'https://company.sharepoint.com/sites/yoursite'\n"
                    "  folder_path: '/sites/yoursite/Shared Documents/folder'\n"
                    "  filename: 'inventory.xlsx'\n"
                    "  client_id: 'your-azure-app-client-id'\n"
                    "  client_secret: 'your-azure-app-client-secret'\n"
                    "  tenant_id: 'your-azure-tenant-id'"
                )
            
            # Validate all required SharePoint fields
            required_fields = ['site_url', 'folder_path', 'filename', 'client_id', 'client_secret', 'tenant_id']
            missing_fields = [
                field for field in required_fields 
                if not self.sharepoint_config.get(field)
            ]

            if missing_fields:
                raise AnsibleParserError(
                    f"Missing required SharePoint configuration fields: {', '.join(missing_fields)}\n"
                    f"All of these fields are required for SharePoint access."
                )

    def _process_local_source(self) -> None:        
        """
        Process Excel file from local filesystem.
        
        Handles the complete workflow for local Excel files:
        1. Extract file path from configuration
        2. Read and validate Excel file
        3. Generate inventory from processed data
        
        This method serves as the coordination point for local file processing,
        delegating specific tasks to specialized methods while handling any
        errors that occur during the process.
        
        Raises:
            AnsibleError: If file reading or inventory generation fails
        """
        try:
            file_path = self.local_config.get('file_path')
            display.vvv(f"Processing local Excel file: {file_path}")
            
            # Read and validate Excel file
            df = self._read_excel_file(file_path)
            
            # Generate inventory from Excel data
            self._populate_inventory(df)
            
        except (AnsibleError, AnsibleParserError):
            raise  # Re-raise known errors without modification
        except Exception as e:
            raise AnsibleError(f"Unexpected error processing local Excel file: {str(e)}")


    def _process_sharepoint_source(self) -> None:
        """
        Process Excel file from SharePoint Online.
        
        Handles the complete workflow for SharePoint Excel files:
        1. Download file from SharePoint using Azure AD authentication
        2. Read and validate the downloaded Excel file
        3. Generate inventory from processed data
        4. Clean up temporary files
        
        This method manages the SharePoint integration complexity while ensuring
        proper cleanup of temporary files even if errors occur during processing.
        
        Raises:
            AnsibleError: If SharePoint access, file processing, or inventory generation fails
        """
        local_file_path = None
        try:
            display.vvv("Processing Excel file from SharePoint Online")
            
            # Download file from SharePoint to temporary location
            local_file_path = self._download_from_sharepoint()
            
            # Read and validate the downloaded Excel file
            df = self._read_excel_file(local_file_path)
            
            # Generate inventory from Excel data
            self._populate_inventory(df)
            
        except (AnsibleError, AnsibleParserError):
            raise  # Re-raise known errors without modification
        except Exception as e:
            raise AnsibleError(f"Unexpected error processing SharePoint Excel file: {str(e)}")
        finally:
            # Always clean up temporary file
            if local_file_path and os.path.exists(local_file_path):
                try:
                    os.remove(local_file_path)
                    display.vvv(f"Cleaned up temporary file: {local_file_path}")
                except OSError as e:
                    display.warning(f"Failed to clean up temporary file {local_file_path}: {str(e)}")


    def _download_from_sharepoint(self) -> str:
        """
        Download Excel file from SharePoint Online using Azure AD authentication.
        
        Establishes a secure connection to SharePoint Online using Azure AD
        application credentials and downloads the specified Excel file to a
        temporary location for processing.
        
        The method uses the Office365-REST-Python-Client library to:
        1. Authenticate using client credentials flow
        2. Connect to the specified SharePoint site
        3. Download the file from the configured folder
        4. Save to a secure temporary location
        
        Authentication requires a registered Azure AD application with:
        - SharePoint permissions (Sites.Read.All or Sites.ReadWrite.All)
        - Client credentials configured (client secret)
        - Proper tenant configuration
        
        Returns:
            str: Path to the downloaded temporary file
            
        Raises:
            AnsibleError: If authentication fails, file not found, or download fails
        """
        config = self.sharepoint_config
        
        try:
            display.vvv("Authenticating with SharePoint Online using Azure AD")
            
            # Configure Azure AD client credentials authentication
            credentials = ClientCredential(
                config['client_id'],
                config['client_secret']
            )
            
            # Establish SharePoint context with authentication
            ctx = ClientContext(config['site_url']).with_credentials(credentials)
            
            # Construct the full file path within SharePoint
            file_url = f"{config['folder_path']}/{config['filename']}"
            display.vvv(f"Target SharePoint file: {file_url}")
            
            # Create secure temporary file for download
            temp_file = tempfile.NamedTemporaryFile(
                delete=False,           # Keep file after closing
                suffix='.xlsx',         # Maintain Excel extension
                prefix='ansible_inventory_'  # Identifiable prefix
            )
            temp_file_path = temp_file.name
            temp_file.close()
            
            display.vvv(f"Downloading to temporary file: {temp_file_path}")
            
            # Download file from SharePoint
            with open(temp_file_path, "wb") as local_file:
                file_obj = ctx.web.get_file_by_server_relative_url(file_url)
                file_obj.download(local_file).execute_query()
            
            display.vvv("SharePoint file downloaded successfully")
            return temp_file_path
            
        except Exception as e:
            raise AnsibleError(f"Error descargando archivo desde SharePoint: {str(e)}")


    def _read_excel_file(self, file_path: str) -> pd.DataFrame:
        """
        Read and validate Excel file using pandas.
        
        Loads the specified Excel file, validates the hostname column exists,
        and optionally filters out rows with empty hostnames.
        
        Args:
            file_path (str): Path to the Excel file
            
        Returns:
            pd.DataFrame: Loaded and validated Excel data
            
        Raises:
            AnsibleParserError: If sheet or hostname column not found
            AnsibleError: If file reading fails
        """
        try:
            # Determine which sheet to read
            sheet_name: Union[str, int] = self.sheet_name or 0
 
            display.vvv(f"Reading Excel file: {file_path}, sheet: {sheet_name}")

            # Load Excel file with pandas
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                engine='openpyxl'
            )

            # Validate hostname column exists
            if self.hostname_column not in df.columns:
                available_columns = ', '.join(df.columns.tolist())
                raise AnsibleParserError(
                    f"Column '{self.hostname_column}' not found in Excel file.\n"
                    f"Available columns: {available_columns}"
                )

            # Filter out empty hostnames if configured
            if self.skip_empty_hosts:
                initial_count = len(df)

                # Remove rows where hostname is null
                df = df.dropna(subset=[self.hostname_column])

                # Remove rows where hostname is empty string after stripping
                df = df[df[self.hostname_column].astype(str).str.strip() != '']

                filtered_count = len(df)

                display.vvv(
                    f"Filtered {initial_count - filtered_count} rows with empty hostnames"
                )
            
            display.vvv(f"Successfully loaded {len(df)} rows, {len(df.columns)} columns")
            return df

        except AnsibleParserError:
            raise # Re-raise configuration errors without modification

        except Exception as e:
            error_msg = str(e)

            # Provide better error messages for common issues
            if "Worksheet named" in error_msg and "not found" in error_msg:
                raise AnsibleParserError(f"Sheet '{self.sheet_name}' not found in Excel file.")
        
            # Generic Excel reading error
            raise AnsibleError(f"Failed to read Excel file: {error_msg}")


    def _populate_inventory(self, df: pd.DataFrame) -> None:
        """
        Generate Ansible inventory from Excel DataFrame.
        
        This is the main inventory generation method that orchestrates the complete
        process of converting Excel data into a structured Ansible inventory with
        hosts, variables, and groups.
        
        The method performs the following operations in sequence:
        1. Apply configured filters to reduce the dataset
        2. Process each remaining row as a host
        3. Create host variables using the compose configuration
        4. Generate automatic groups using keyed_groups
        5. Create conditional groups using group expressions
        
        Args:
            df (pd.DataFrame): Validated Excel data to process
            
        Note:
            If no hosts remain after filtering, the method logs a warning
            but continues successfully (empty inventory is valid).
        """
        # Apply configured filters to reduce dataset
        df = self._apply_filters(df)
        
        # Check if any hosts remain after filtering
        if df.empty:
            display.warning(
                "No hosts found after applying filters. "
                "Inventory will be empty. Check your filter configuration."
            )
            return
        
        display.vvv(f"Generating inventory for {len(df)} hosts")
        
        # Process each row as an inventory host
        for _, row in df.iterrows():
            hostname = str(row[self.hostname_column]).strip()
            
            # Skip rows with empty hostnames (shouldn't happen due to earlier filtering)
            if not hostname:
                continue
            
            # Add host to Ansible inventory
            self.inventory.add_host(hostname)
            
            # Apply compose with intelligent variable filtering
            if self.compose:
                # Filter only valid values using dictionary comprehension
                # This automatically removes NaN, None, empty strings, and null values
                host_vars = {
                    column: value 
                    for column, value in row.to_dict().items() 
                    if self._is_valid_value(value)
                }
                
                # Only apply compose if there are valid variables to map
                if host_vars:
                    self._set_composite_vars(self.compose, host_vars, hostname)
                else:
                    display.vvv(
                        f"Host '{hostname}': no valid variables found after filtering, "
                        f"only hostname will be available"
                    )
        
        # Create automatic groups based on column values
        if self.keyed_groups:
            for keyed_group in self.keyed_groups:
                if isinstance(keyed_group, dict) and 'key' in keyed_group:
                    # Advanced keyed group configuration
                    self._create_keyed_groups(df, keyed_group)
                elif isinstance(keyed_group, str):
                    # Simple keyed group (just column name)
                    self._create_keyed_groups(df, {'key': keyed_group})
        
        # Create conditional groups based on expressions
        if self.groups:
            self._create_conditional_groups(df)


    def _apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Apply configured filters to the DataFrame.
        
        Filters are applied with AND logic - all filters must match for a host
        to be included. Missing columns generate warnings but don't cause failures.
        
        Args:
            df (pd.DataFrame): Source data to filter
            
        Returns:
            pd.DataFrame: Filtered data containing only matching rows
        """
        if not self.filters:
            return df
        
        display.vvv(f"Applying filters: {self.filters}")
        initial_count = len(df)
        
        for column, value in self.filters.items():
            # Skip empty or None filter values
            if value is not None and value != '':
                if column not in df.columns:
                    display.warning(
                        f"Filter column '{column}' not found in Excel, ignoring filter"
                    )
                    continue
                
                # Apply string-based equality filter
                df = df[df[column].astype(str).str.strip() == str(value).strip()]
                display.vvv(f"Filtro {column}={value} aplicado, {len(df)} filas restantes")
        
        filtered_count = len(df)
        display.vvv(f"Filtros aplicados: {initial_count} -> {filtered_count} filas")
        return df


    def _is_valid_value(self, value: Any) -> bool:
        """
        Check if a value is valid for adding as an inventory variable.

        Filters out empty/null values commonly found in Excel files:
        - pandas NaN (empty cells)
        - Python None
        - Empty strings or whitespace-only strings
        - String representations of null ('nan', 'none', 'null', 'n/a')
        
        Args:
            value (Any): Value to check from Excel data

        Returns:
            bool: True if value should be included, False if it should be skipped
        """
        # Check for pandas NaN (empty Excel cells)
        if pd.isna(value):
            return False
        
        # Check for Python None
        if value is None:
            return False

        # Convert to string for additional checks
        str_value = str(value).strip().lower()

        # Check for empty string after trimming whitespace
        if str_value == '':
            return False

        # Check for explicit null representations
        null_representations = ['nan', 'none', 'null', 'n/a', 'na', '#n/a', '#null']
        if str_value in null_representations:
            return False

        # Value is valid
        return True


    def _create_keyed_groups(self, df: pd.DataFrame, keyed_group_config: Dict[str, Any]) -> None:
        """
        Create inventory groups based on unique values in a column.
        
        Automatically creates one inventory group for each unique value found
        in the specified column. Supports optional prefixes and custom separators.
        
        Args:
            df (pd.DataFrame): Source data for grouping
            keyed_group_config (Dict[str, Any]): Group configuration containing:
                - key (str): Column name to group by
                - prefix (str, optional): Prefix for group names
                - separator (str, optional): Separator between prefix and value
                
        Note:
            Group names are automatically normalized (lowercase, spaces to underscores)
        """
        key_column = keyed_group_config['key']
        prefix = keyed_group_config.get('prefix', '')
        separator = keyed_group_config.get('separator', '_')
        
        if key_column not in df.columns:
            display.warning(f"Keyed group column '{key_column}' not found in Excel")
            return
        
        display.vvv(
            f"Creating groups by column '{key_column}'"
            + (f" with prefix '{prefix}'" if prefix else "")
        )
        
        # Group DataFrame by the specified colu
        for group_value, group_df in df.groupby(key_column):
            # Skip null or empty values
            if pd.notna(group_value) and str(group_value).strip():
                # Normalize group name
                base_name = str(group_value).lower().replace(' ', '_').replace('-', '_')
                
                # Apply prefix if specified
                group_name = f"{prefix}{separator}{base_name}" if prefix else base_name
                
                # Create inventory group
                self.inventory.add_group(group_name)
                
                # Add all hosts in this group
                for _, row in group_df.iterrows():
                    hostname = str(row[self.hostname_column]).strip()
                    if hostname:
                        self.inventory.add_child(group_name, hostname)
                
                display.vvv(f"Created group '{group_name}' with {len(group_df)} hosts")


    def _create_conditional_groups(self, df: pd.DataFrame) -> None:
        """
        Create inventory groups based on conditional expressions.
        
        Evaluates Python-like expressions for each host to determine group
        membership. This enables complex grouping logic like combining multiple
        criteria or using comparison operators.
        
        Expression features:
        - Comparison operators: ==, !=, <, <=, >, >=
        - Logical operators: and, or, not
        - Membership operators: in, not in
        - Variable names must match Excel column names exactly
        - Safe evaluation prevents code injection
        
        Args:
            df (pd.DataFrame): Source data for conditional evaluation
            
        Example expressions:
            "Environment == 'Production'"
            "CPU_Count >= 8 and Memory_GB >= 32"
            "Environment == 'Production' and Role in ['Database', 'Web']"
        """
        display.vvv(f"Creating {len(self.groups)} conditional groups")
        
        for group_name, condition in self.groups.items():
            try:
                # Create the inventory group (even if it ends up empty)
                self.inventory.add_group(group_name)
                hosts_added = 0
                
                # Evaluate the condition for each host
                for _, row in df.iterrows():
                    hostname = str(row[self.hostname_column]).strip()
                    if not hostname:
                        continue
                    
                    # Create evaluation context from Excel row data
                    context = row.to_dict()
                    
                    # Safely evaluate the conditional expression
                    if self._evaluate_condition(condition, context):
                        self.inventory.add_child(group_name, hostname)
                        hosts_added += 1
                
                display.vvv(f"Conditional group '{group_name}': {hosts_added} hosts matched")
                
            except Exception as e:
                display.warning(
                    f"Error creating conditional group '{group_name}' with condition '{condition}': {str(e)}"
                )

    
    def _evaluate_condition(self, condition: str, context: Dict[str, Any]) -> bool:
        """
        Safely evaluate a Python-like expression for conditional grouping.
        
        Uses Abstract Syntax Tree (AST) parsing to safely evaluate expressions
        without the security risks of eval(). Only allows safe operators and
        prevents execution of arbitrary code.
        
        Security features:
        - No function calls allowed
        - No imports or attribute access
        - Only whitelisted operators permitted
        - Variables resolved from provided context only
        
        Supported operators:
        - Comparison: ==, !=, <, <=, >, >=
        - Logical: and, or, not
        - Membership: in, not in
        - Literals: strings, numbers, booleans
        
        Args:
            condition (str): Python-like expression to evaluate
            context (Dict[str, Any]): Variable context (Excel row data)
            
        Returns:
            bool: Result of expression evaluation
            
        Example:
            condition = "Environment == 'Production' and CPU_Count >= 8"
            context = {'Environment': 'Production', 'CPU_Count': 16}
            → Returns True
        """
        # Define allowed operators for security (whitelist approach)
        safe_operators = {
            ast.Eq: operator.eq,        # ==
            ast.NotEq: operator.ne,     # !=
            ast.Lt: operator.lt,        # <
            ast.LtE: operator.le,       # <=
            ast.Gt: operator.gt,        # >
            ast.GtE: operator.ge,       # >=
            ast.And: operator.and_,     # and
            ast.Or: operator.or_,       # or
            ast.Not: operator.not_,     # not
            ast.In: lambda x, y: x in y,        # in
            ast.NotIn: lambda x, y: x not in y  # not in
        }
        
        def safe_eval(node: ast.AST, context: Dict[str, Any]) -> Any:
            """
            Recursively and safely evaluate AST nodes.
            
            Args:
                node (ast.AST): AST node to evaluate
                context (Dict[str, Any]): Variable resolution context
                
            Returns:
                Any: Evaluated value
                
            Raises:
                ValueError: If unsafe operations are attempted
            """
            # Handle literal values
            if isinstance(node, ast.Constant):  # Python 3.8+
                return node.value
            elif isinstance(node, ast.Str):  # Python < 3.8 string literals
                return node.s
            elif isinstance(node, ast.Num):  # Python < 3.8 numeric literals
                return node.n
            
            # Handle variable references
            elif isinstance(node, ast.Name):
                var_name = node.id
                if var_name in context:
                    return context[var_name]
                else:
                    raise ValueError(f"Variable '{var_name}' not found in context")
            
            # Handle comparison operations (==, !=, <, etc.)
            elif isinstance(node, ast.Compare):
                left = safe_eval(node.left, context)
                # Handle chained comparisons (a < b < c)
                for op, comparator in zip(node.ops, node.comparators):
                    right = safe_eval(comparator, context)
                    if type(op) not in safe_operators:
                        raise ValueError(f"Unsafe comparison operator: {type(op).__name__}")
                    if not safe_operators[type(op)](left, right):
                        return False
                    left = right  # For chained comparisons
                return True
            
            # Handle boolean operations (and, or)
            elif isinstance(node, ast.BoolOp):
                if type(node.op) not in safe_operators:
                    raise ValueError(f"Unsafe boolean operator: {type(node.op).__name__}")
                
                if isinstance(node.op, ast.And):
                    return all(safe_eval(value, context) for value in node.values)
                elif isinstance(node.op, ast.Or):
                    return any(safe_eval(value, context) for value in node.values)
            
            # Handle unary operations (not)
            elif isinstance(node, ast.UnaryOp):
                if type(node.op) not in safe_operators:
                    raise ValueError(f"Unsafe unary operator: {type(node.op).__name__}")
                return safe_operators[type(node.op)](safe_eval(node.operand, context))
            
            # Handle list literals for 'in' operations
            elif isinstance(node, ast.List):
                return [safe_eval(elem, context) for elem in node.elts]
            
            # Reject any other node types as unsafe
            else:
                raise ValueError(f"Unsafe expression type: {type(node).__name__}")
        
        try:
            # Clean and normalize the evaluation context
            clean_context = {}
            for key, value in context.items():
                # Only include non-null values to avoid pandas NaN issues
                if pd.notna(value):
                    clean_key = str(key).strip()
                    clean_context[clean_key] = value
            
            # Parse the expression into an Abstract Syntax Tree
            tree = ast.parse(condition, mode='eval')
            
            # Safely evaluate the expression
            result = safe_eval(tree.body, clean_context)
            
            # Log detailed evaluation for debugging (very verbose mode)
            display.vvvv(f"Condition '{condition}' evaluated to: {result}")
            
            return bool(result)
            
        except Exception as e:
            # Log evaluation errors but don't fail the entire process
            display.warning(f"Error evaluating condition '{condition}': {str(e)}")
            return False
