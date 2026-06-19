"""
FLOOR_EMOJI.py — Single source of truth for floor → emoji mapping.

Import from here in ZONE_PICKER.py, inject_backend.py, PDF build scripts,
and anywhere else that shows floor icons. Changing emojis in ONE place
(here) propagates to every consumer automatically.

History:
- Created 2026-04-09 after repeated drift between ZONE_PICKER.py,
  inject_backend.py, and ephemeral PDF build scripts.
- Replaces local FLOOR_EMOJI dicts that had all gotten out of sync.

Usage:
    from FLOOR_EMOJI import FLOOR_EMOJI, DISPLAY_ORDER
    emoji = FLOOR_EMOJI.get(floor, '📍')  # fallback for unknown floors
"""

FLOOR_EMOJI = {
    'Upstairs':    '🧺',
    'Main Floor':  '🧺',
    'Basement':    '🧺',
    'Digital':     '💻',
    'Plant':       '🌱',
    'Personal':    '👱‍♀️',
    'Maintenance': '🏠',
    'Business':    '💸',
    'Frog':        '🐸',
}

DISPLAY_ORDER = [
    'Upstairs',
    'Main Floor',
    'Basement',
    'Digital',
    'Plant',
    'Personal',
    'Maintenance',
    'Business',
    'Frog',
]
