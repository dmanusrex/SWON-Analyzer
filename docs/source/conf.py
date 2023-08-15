# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information
import sphinx_rtd_theme

project = 'SWON Analyzer'
copyright = '2023, Darren Richer'
author = 'Darren Richer'
# release = '0.1'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = [
    # https://plantweb.readthedocs.io/#sphinx-directives
    # https://plantweb.readthedocs.io/examples.html
    "plantweb.directive",
    "sphinx_rtd_theme",
    # https://sphinx-tabs.readthedocs.io/en/latest/
    "sphinx_tabs.tabs",
]
templates_path = ['_templates']
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store", ".venv", "README.md"]



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'sphinx_rtd_theme'
html_static_path = ['_static']

# Logo on the sidebar
html_logo = "media/swon-logo.png"

# Favicon
html_favicon = "media/swon-logo.ico"
