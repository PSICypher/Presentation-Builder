"""Schema loader — YAML serialization and deserialization for TemplateSchema.

Provides round-trip save/load so schemas can be reviewed, version-controlled,
and edited as human-readable YAML configuration files.
"""

from pathlib import Path

import yaml

from .models import TemplateSchema


def save_schema(schema: TemplateSchema, path: str | Path) -> None:
    """Serialize a TemplateSchema to a YAML file."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = schema.to_dict()
    with open(path, "w") as f:
        yaml.dump(data, f, default_flow_style=False, sort_keys=False,
                  allow_unicode=True, width=120)


def load_schema(path: str | Path) -> TemplateSchema:
    """Deserialize a TemplateSchema from a YAML file."""
    path = Path(path)
    with open(path) as f:
        data = yaml.safe_load(f)
    return TemplateSchema.from_dict(data)
