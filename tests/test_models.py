from __future__ import annotations

import unittest

from pops_ingest.models import Bounds


class BoundsTests(unittest.TestCase):
    def test_round_trip_and_iou(self) -> None:
        first = Bounds.from_a1("B3:F10")
        second = Bounds.from_a1("E8:H12")
        self.assertEqual(first.ref, "B3:F10")
        self.assertEqual(first.width, 5)
        self.assertEqual(first.height, 8)
        self.assertEqual(first.intersection_area(second), 6)
        self.assertGreater(first.iou(second), 0)
        self.assertEqual(first.union(second).ref, "B3:H12")


if __name__ == "__main__":
    unittest.main()

