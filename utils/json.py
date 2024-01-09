"""utils.json

Serialize and deserialize `json` data.

Functions:
- cls_to_json(instance, file_name):
    - Generate `json` file from class instance.
- json_to_cls(cls, file_path):
    - Generate `instance` from json file at `file_name`
"""
import json
from pathlib import Path
import numpy as np
from utils.file_excel import XW

def _serialize(data, black_list=None):
    """Serializes `data` input to comply with JSON formatting."""
    if black_list is None:
        black_list = []
    if isinstance(data, np.ndarray):
        return {
            '__np__': True, 
            'data': data.tolist(),
            'dtype': str(data.dtype)
        }
    if isinstance(data, dict):
        return {
            key: None if isinstance(value, float) and np.isnan(value) else _serialize(value, black_list) 
            for key, value in data.items() 
            if key not in black_list and not key.startswith('_')
        }
    if isinstance(data, (list, tuple)):
        return [
            None if isinstance(x, float) and np.isnan(x) else _serialize(x, black_list) 
            for x in data
        ]
    if isinstance(data, Path):
        return f"{data.parts[-3]}/{data.parts[-2]}/{data.parts[-1]}"
    if isinstance(data, XW):
        ranges = data.ranges
        return {
            '__xw__': True,
            'book': data.book.fullname,
            'sheet_name': data.sheet.name,
            'ranges': [ranges[key].name.name for key in ranges.keys()],
            'attr_names': list(ranges.keys()),
        }
    return data

def _deserialize(data):
    """Deserializes `data` to reconstruct the instance."""
    if isinstance(data, dict):
        if data.get('__np__', False):
            return np.array(data['data'], dtype=data['dtype'])
        if isinstance(data.get('data', None), dict):
            return data['data']
        if data.get('__xw__', False):
            return XW(
                book=Path(data['book']), sheet_name=data['sheet_name'], ranges=data['ranges'], attr_names=data['attr_names']
                )
        return {
            key: _deserialize(value) for key, value in data.items()
        }
    return [_deserialize(x) for x in data] if isinstance(data, list) else data

def cls_to_json(instance, file_name:str or Path, indent=4, black_list:list[str]=None)->None:
    """Generate `json` file from class instance.
    
    Parameters:
    - instance: class instance
    - file_name (str or Path): Path to save json file.
    - indent (int, optional): Indent level for json file (usually 2 or 4), default=4.
    - black_list (list[str], optional): List of blacklisted attributes.
    """
    with open(file_name, 'w', encoding='utf-8') as f:
        data = _serialize(instance.__dict__, black_list)
        json.dump(data, f, indent=indent)

def json_to_cls(cls, file_name:str or Path):
    """Generate `instance` from json file at `file_name`.
    
    Parameters:
    - cls : Class to instantiate.
    - file_name (str or Path): Path to json file.
    
    Returns:
    - instance: class instance
    """
    with open(file_name, 'r') as f:
        data = json.load(f)
    data = _deserialize(data)
    instance = cls.__new__(cls)
    instance.__dict__ = data
    return instance
