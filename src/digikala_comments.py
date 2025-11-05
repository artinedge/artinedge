"""Extract Digikala product comments and store them in an Excel file.

This module provides a small command line interface that can be used to
retrieve all comments that are publicly visible for a Digikala product. The
script expects either a Digikala product URL (``https://www.digikala.com/...``),
the ``dkp-<id>`` identifier fragment or the plain numeric product ID. The
retrieved comments are normalised and saved in an Excel spreadsheet.

Example
-------

Run the script from the command line to fetch the comments for a specific
product and store them in ``comments.xlsx``::

    python src/digikala_comments.py "https://www.digikala.com/product/dkp-7068663" --output comments.xlsx

The Digikala API requires an ordinary browser user-agent header. If you hit
rate limits you can adjust the request delay via ``--delay``.
"""

from __future__ import annotations

import argparse
import dataclasses
import re
import time
from typing import Iterable, Iterator, List, Optional

try:
    import requests
except ModuleNotFoundError as exc:  # pragma: no cover - this path is for runtime UX only
    raise SystemExit(
        "The 'requests' package is required to run this script. Install it via 'pip install -r requirements.txt'."
    ) from exc

try:
    import pandas as pd
except ModuleNotFoundError as exc:  # pragma: no cover - this path is for runtime UX only
    raise SystemExit(
        "The 'pandas' package is required to run this script. Install it via 'pip install -r requirements.txt'."
    ) from exc

API_URL_TEMPLATE = "https://api.digikala.com/v1/product/{product_id}/comments/"
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/115.0 Safari/537.36"
    ),
    "Accept": "application/json",
}
ID_PATTERN = re.compile(r"(?:dkp-)?(\d+)")


@dataclasses.dataclass
class Comment:
    """Normalised representation of a Digikala comment."""

    comment_id: Optional[int]
    author: Optional[str]
    title: Optional[str]
    body: Optional[str]
    rating: Optional[float]
    recommendation_status: Optional[str]
    created_at: Optional[str]
    purchase_status: Optional[str]
    positive_points: Optional[str]
    negative_points: Optional[str]
    likes: Optional[int]
    dislikes: Optional[int]

    @classmethod
    def from_api(cls, payload: dict) -> "Comment":
        """Create a :class:`Comment` from a raw API payload.

        The API structure used by Digikala occasionally changes. The logic below
        focuses on widely observed fields and falls back gracefully when a
        specific key is missing.
        """

        def _get(*keys, default=None):
            value = payload
            for key in keys:
                if not isinstance(value, dict) or key not in value:
                    return default
                value = value[key]
            return value

        def _join_points(points: Optional[Iterable[str]]) -> Optional[str]:
            if points is None:
                return None
            if isinstance(points, str):
                return points
            if isinstance(points, Iterable):
                cleaned = [p.strip() for p in points if p]
                return "\n".join(cleaned) if cleaned else None
            return None

        author = _get("author") or _get("user", "name") or _get("user", "username")
        purchase_status = _get("purchase_status")
        if purchase_status is None:
            if _get("is_buyer"):
                purchase_status = "buyer"
            elif _get("has_bought"):
                purchase_status = "buyer"

        return cls(
            comment_id=_get("id") or _get("comment_id"),
            author=author,
            title=_get("title"),
            body=_get("body") or _get("text") or _get("content"),
            rating=_get("rate") or _get("rating"),
            recommendation_status=_get("recommendation_status") or _get("recommendation"),
            created_at=_get("created_at") or _get("date") or _get("created_on"),
            purchase_status=purchase_status,
            positive_points=_join_points(_get("advantages") or _get("positives")),
            negative_points=_join_points(_get("disadvantages") or _get("negatives")),
            likes=_get("likes") or _get("like_count"),
            dislikes=_get("dislikes") or _get("dislike_count"),
        )


class DigikalaCommentClient:
    """Client responsible for fetching comment pages from Digikala."""

    def __init__(self, product_id: str, delay: float = 0.5, session: Optional[requests.Session] = None):
        self.product_id = product_id
        self.delay = max(0.0, delay)
        self.session = session or requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)

    def fetch_page(self, page: int) -> dict:
        url = API_URL_TEMPLATE.format(product_id=self.product_id)
        response = self.session.get(url, params={"page": page}, timeout=20)
        response.raise_for_status()
        return response.json()

    def iter_comments(self) -> Iterator[Comment]:
        page = 1
        while True:
            data = self.fetch_page(page)
            comments_section = self._extract_comment_section(data)
            items = comments_section.get("items") or comments_section.get("data") or []
            if not items:
                break

            for item in items:
                if isinstance(item, dict):
                    yield Comment.from_api(item)

            if not self._has_next_page(comments_section):
                break

            page += 1
            if self.delay:
                time.sleep(self.delay)

    @staticmethod
    def _extract_comment_section(payload: dict) -> dict:
        containers: List[dict] = []
        data = payload.get("data") if isinstance(payload, dict) else None
        if isinstance(data, dict):
            for key in ("comments", "comment", "reviews"):
                section = data.get(key)
                if isinstance(section, dict):
                    containers.append(section)
        if not containers and isinstance(payload, dict):
            for key in ("comments", "data"):
                section = payload.get(key)
                if isinstance(section, dict):
                    containers.append(section)

        return containers[0] if containers else {}

    @staticmethod
    def _has_next_page(section: dict) -> bool:
        paging = section.get("paging") or {}
        if isinstance(paging, dict):
            if paging.get("next"):
                return True
            next_page = paging.get("next_page")
            if isinstance(next_page, int) and next_page > 0:
                return True
            total_pages = paging.get("total_pages")
            current_page = paging.get("current_page") or paging.get("page")
            if isinstance(total_pages, int) and isinstance(current_page, int) and current_page < total_pages:
                return True

        links = section.get("links") or {}
        if isinstance(links, dict) and links.get("next"):
            return True

        return False


def extract_product_id(value: str) -> str:
    """Extract the numeric product ID from a Digikala URL or identifier."""

    match = ID_PATTERN.search(value)
    if not match:
        raise ValueError(
            "Unable to locate a Digikala product identifier. Provide the product URL or the 'dkp-<id>' value."
        )
    return match.group(1)


def comments_to_dataframe(comments: Iterable[Comment]) -> pd.DataFrame:
    records = [dataclasses.asdict(comment) for comment in comments]
    if not records:
        return pd.DataFrame(columns=[field.name for field in dataclasses.fields(Comment)])
    return pd.DataFrame.from_records(records)


def save_to_excel(df: pd.DataFrame, output_path: str) -> None:
    df.to_excel(output_path, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("product", help="Digikala product URL or identifier (e.g. https://www.digikala.com/product/dkp-12345/)")
    parser.add_argument(
        "--output",
        "-o",
        default="digikala_comments.xlsx",
        help="Path to the Excel file that should be created (default: digikala_comments.xlsx)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.5,
        help="Delay in seconds between API requests (default: 0.5)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    product_id = extract_product_id(args.product)

    client = DigikalaCommentClient(product_id, delay=args.delay)
    comments = list(client.iter_comments())
    df = comments_to_dataframe(comments)
    save_to_excel(df, args.output)

    print(f"Saved {len(df)} comments to '{args.output}'.")


if __name__ == "__main__":  # pragma: no cover - entry point
    main()
