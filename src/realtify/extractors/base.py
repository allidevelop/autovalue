from __future__ import annotations

from abc import ABC, abstractmethod

from realtify.models import Comparable, PropertyType, TransactionType
from realtify.screenshot_sources import PageSnapshot
from realtify.source_config import SourceDefinition


class ListingExtractor(ABC):
    @abstractmethod
    def extract(
        self,
        snapshot: PageSnapshot,
        source: SourceDefinition,
        *,
        property_type: PropertyType,
        transaction_type: TransactionType,
    ) -> Comparable:
        raise NotImplementedError

