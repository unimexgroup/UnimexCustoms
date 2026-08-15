"""
Single source of truth for the tool's version.

Bump this constant, add a CHANGELOG entry, commit, then push a matching tag
(``customs-vX.Y.Z``). The GitHub Actions release workflow parses this file and
refuses to build if the tag's number disagrees with the constant here.
"""

CUSTOMS_VERSION = "1.3.0"
