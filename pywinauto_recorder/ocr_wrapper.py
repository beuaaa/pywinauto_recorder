"""
Placeholder of the OCR_Text elements: the optical character recognition (OCR) is disabled.

The former implementation, based on EasyOCR, is in the history of the Git repository (version 0.6.9 of this file).
"""


class OCRWrapper:
	"""Wrapper of a text found by the OCR: there is no such text, because the OCR is disabled."""


def find_all_ocr(
		wrapper, allowlist=None, mag_ratio=2, width_ths=0.5, contrast_ths=0.1, adjust_contrast=0.5, rotation_info=None):
	"""Returns the texts found by the OCR in the element: none, because the OCR is disabled."""
	return []
