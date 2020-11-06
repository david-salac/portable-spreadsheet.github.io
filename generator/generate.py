from pathlib import Path
import shutil
from typing import List, Union
import pkgutil

import crinita as cr

# Iterate through all entities
ENTITIES: List[Union[cr.Page, cr.Article]] = []
for single_file in pkgutil.walk_packages(['.']):
    try:
        if single_file.name == 'generate':
            continue
        pack = __import__(single_file.name)
        ENTITIES.append(pack.ENTITY)
    except:  # noqa
        continue

sites = cr.Sites(ENTITIES)
# ========= CONFIGURATION =========
# Path to outputs
output_directory: Path = Path('../docs/')
# Resource directory
resource_directory: Path = Path('RESOURCES')

# Add template path:
cr.Config.templates_path = Path('templates')
# Configure blog name
cr.Config.site_logo_text = "Portable Spreadsheet"
cr.Config.site_title = "Portable Spreadsheet"
cr.Config.append_to_menu = (
    {
        'title': "GitHub Project",
        'url': "https://github.com/david-salac/Portable-spreadsheet-generator",
        'menu_position': 30
    },
)
cr.Config.text_sections_in_right_menu = (
    {
        "header": "Portable Spreadsheet",
        "content": f'Simple spreadsheet that keeps tracks of each operation in defined languages. Logic allows export sheets to Excel files (and see how each cell is computed), to the JSON strings with description of computation e. g. in native language. Other formats like HTML, CSV and Markdown (MD) are also supported.<p>Generated using <a href="http://www.crinita.com/">Crinita</a> version {cr.__version__}</p>'
    },)
cr.Config.default_meta_description = "Spreadsheet generator that keeps tracks of each operation in defined languages. Allows export sheets to Excel files (and see how all cells are computed)."
cr.Config.default_meta_meta_author = "Portable Spreadsheet team"
cr.Config.default_meta_keywords = "spreadsheet, Excel, JSON, keeping track, operations, cells"
cr.Config.site_home_url = "/"
cr.Config.site_map_url_prefix = "https://portable-spreadsheet.com/"
cr.Config.robots_txt = """User-agent: *
Allow: /
Sitemap: https://portable-spreadsheet.com/sitemap.xml"""
# =================================

# Remove existing content
if output_directory.exists():
    shutil.rmtree(output_directory)

# Create it de novo
output_directory.mkdir()

# Generate sites
sites.generate_pages(output_directory)

# Add resources
shutil.copytree(resource_directory, output_directory, dirs_exist_ok=True)
