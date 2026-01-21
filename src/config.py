"""
Configuration management module for Rubick CAST Formatting.

Provides centralized configuration with environment variable support
and Pydantic validation for type safety.
"""

from pathlib import Path
from typing import Optional
from functools import lru_cache

from pydantic import BaseModel, Field
from pydantic_settings import BaseSettings


class AppConfig(BaseModel):
    """Application configuration model."""
    app_name: str = Field(default="Rubick CAST Formatting", description="Application display name")
    app_version: str = Field(default="0.1.0", description="Application version")
    debug_mode: bool = Field(default=False, description="Enable debug logging")
    max_file_size_mb: int = Field(default=50, description="Maximum upload file size in MB")
    temp_dir: str = Field(default="/tmp/rubick-cast", description="Temporary directory for processing")


class PathsConfig(BaseModel):
    """Path configuration model."""
    templates_dir: Path = Field(default=Path("templates"), description="Directory for template files")
    output_dir: Path = Field(default=Path("output"), description="Directory for output files")
    temp_dir: Path = Field(default=Path("/tmp/rubick-cast"), description="Temporary processing directory")

    def ensure_directories(self) -> None:
        """Create output directories if they don't exist."""
        self.templates_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(parents=True, exist_ok=True)


class ExportConfig(BaseModel):
    """Export configuration model."""
    default_template: str = Field(default="", description="Default template filename")
    output_prefix: str = Field(default="Formatted_", description="Prefix for output filenames")
    preserve_unknown_columns: bool = Field(default=False, description="Keep columns not in template")


class ImportConfig(BaseModel):
    """Import configuration model."""
    size_chart_filename: str = Field(default="sku.xlsx", description="Default size chart filename")
    product_details_filename: str = Field(default="style.xlsx", description="Default product details filename")
    output_filename: str = Field(default="Batch_Merged_With_Types_Values.xlsx", description="Default output filename")
    exclude_sheets: list[str] = Field(default=["masterdata"], description="Sheet names to exclude")
    style_column_patterns: list[str] = Field(
        default=["style[_-]?id", "sku", "styleid"],
        description="Regex patterns for style ID column detection"
    )
    image_column_patterns: list[str] = Field(
        default=["image", "img", "url", "cdn"],
        description="Regex patterns for image column detection"
    )


class Settings(BaseSettings):
    """Main settings class combining all configuration sections."""
    app: AppConfig = Field(default_factory=AppConfig)
    paths: PathsConfig = Field(default_factory=PathsConfig)
    export: ExportConfig = Field(default_factory=ExportConfig)
    import_config: ImportConfig = Field(default_factory=ImportConfig)

    class Config:
        env_prefix = "RUBICK_CAST_"
        env_nested_delimiter = "__"

    @classmethod
    def from_env(cls) -> "Settings":
        """Load settings from environment variables."""
        return cls()


@lru_cache()
def get_settings() -> Settings:
    """Get cached settings instance."""
    return Settings.from_env()


def reload_settings() -> Settings:
    """Clear cache and reload settings."""
    get_settings.cache_clear()
    return get_settings()
