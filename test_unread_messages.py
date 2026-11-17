import unittest
from datetime import datetime, timezone

from unread_messages import Message, demo_source, get_all_unread


class GetAllUnreadTest(unittest.TestCase):
    def test_merges_sources_newest_first(self):
        result = get_all_unread([demo_source])
        self.assertEqual(len(result), 3)
        dates = [m.received for m in result]
        self.assertEqual(dates, sorted(dates, reverse=True))

    def test_broken_source_does_not_hide_others(self):
        def broken():
            raise ConnectionError("немає мережі")

        extra = Message("Viber", "Іра", "Привіт", datetime(2026, 1, 1, tzinfo=timezone.utc))
        result = get_all_unread([broken, lambda: [extra]])
        self.assertEqual(result, [extra])

    def test_no_sources(self):
        self.assertEqual(get_all_unread([]), [])


if __name__ == "__main__":
    unittest.main()
