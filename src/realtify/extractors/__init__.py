from __future__ import annotations

from realtify.extractors.base import ListingExtractor
from realtify.extractors.dimria import DimriaListingExtractor
from realtify.extractors.generic import GenericListingExtractor
from realtify.extractors.rieltor import RieltorListingExtractor


def extractor_for_source(source_name: str) -> ListingExtractor:
    if source_name == "rieltor":
        return RieltorListingExtractor()
    if source_name == "dimria":
        return DimriaListingExtractor()
    return GenericListingExtractor(source_name=source_name)


__all__ = [
    "DimriaListingExtractor",
    "GenericListingExtractor",
    "ListingExtractor",
    "RieltorListingExtractor",
    "extractor_for_source",
]
