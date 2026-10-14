"""OISI Component 4 - shared extraction rules (bots, identities, window).

Every script imports these definitions so that ALL repositories are processed
with exactly the same rules. Change them here, never inside a script.
"""
import re
import unicodedata
from datetime import datetime, timezone

# ---------------------------------------------------------------------------
# 1. Observation window (start inclusive, end exclusive, UTC)
# ---------------------------------------------------------------------------
WINDOW_START = datetime(2025, 10, 1, tzinfo=timezone.utc)
WINDOW_END = datetime(2026, 10, 1, tzinfo=timezone.utc)
S = WINDOW_START.timestamp()
E = WINDOW_END.timestamp()
MID = (S + E) / 2  # split for contributor retention (two half-years)
DATE_FIELD = "author"  # commits are dated by author date (%at)

# Cohort cut-offs: an item is eligible only if its full follow-up period
# ends before WINDOW_END, so no item is censored.
ISSUE_RESPONSE_DAYS = 14
PR_RESPONSE_DAYS = 7
ISSUE_CLOSE_DAYS = 90
PR_DECISION_DAYS = 30
MIN_ITEMS = 10  # metrics 8-11 need >= 10 eligible items, else "insufficient"
CLOSE_COMMENT_GRACE_S = 60  # a comment by the closer within 60 s of closing = the close itself

# ---------------------------------------------------------------------------
# 2. Bot rules
# ---------------------------------------------------------------------------
# A. Explicit bot accounts / identities (lower-case). Matched against GitHub
#    login, git author name and git author email local part.
BOT_ACCOUNTS = {
    "dependabot", "dependabot[bot]", "dependabot-preview[bot]", "renovate[bot]",
    "renovate-bot", "github-actions", "github-actions[bot]", "pre-commit-ci[bot]",
    "weblate", "snyk-bot", "greenkeeper[bot]", "semantic-release-bot",
    "actions-user", "codecov[bot]", "allcontributors[bot]", "mergify[bot]",
    "imgbot[bot]", "stale[bot]", "copilot", "copilot-swe-agent[bot]",
    "coderabbitai[bot]", "sonarcloud[bot]", "netlify[bot]", "vercel[bot]",
    "deepsource-autofix[bot]", "transifex-integration[bot]", "crowdin-bot",
    "lgtm-com[bot]", "gitlab-bot", "translate-bot", "release-please[bot]",
}
# B. Generic patterns (case-insensitive) on "name email" or login.
BOT_PATTERNS = re.compile(
    r"\[bot\]|dependabot|renovate|github-actions|pre-commit|weblate|snyk|"
    r"greenkeeper|semantic-release|actions-user|\bbot@|-bot\b|\bbot\b|"
    r"noreply@weblate|transifex|crowdin|translations?@",
    re.I,
)
# GitHub API: user.type == "Bot" is always a bot (authoritative).


def is_bot(name: str = "", email: str = "", login: str = "", user_type: str = "") -> bool:
    if user_type and user_type.lower() == "bot":
        return True
    for v in (login, name, (email or "").split("@")[0]):
        if v and v.lower() in BOT_ACCOUNTS:
            return True
    return bool(BOT_PATTERNS.search(f"{name} {email} {login}"))


# ---------------------------------------------------------------------------
# 3. Identity resolution
# ---------------------------------------------------------------------------
# Applied in this order inside ONE repository (identities are never merged
# across repositories):
#   R1  .mailmap in the repository (git applies it via %aN / %aE).
#   R2  identical e-mail address (case-insensitive) = same person.
#   R3  GitHub no-reply e-mail "12345+login@users.noreply.github.com" or
#       "login@users.noreply.github.com" -> keyed by the GitHub login.
#   R4  GitHub API commit.author.login (when the extractor has it) is
#       authoritative: identical login = same person.
#   R5  identical normalised FULL name (>= 2 words, not generic) = same
#       person, flagged "name-match" so it can be audited.
#   R6  manual map file identity_map.csv (repo,email,canonical_id,evidence).
# Never merge on similar-but-different names, first names only, or initials.
NOREPLY = re.compile(r"^(?:\d+\+)?([^@]+)@users\.noreply\.github\.com$", re.I)
GENERIC_NAMES = {"root", "admin", "ubuntu", "user", "unknown", "developer",
                 "dev", "test", "your name", "administrator", "system"}


def norm_name(name: str) -> str:
    n = unicodedata.normalize("NFKD", name or "").encode("ascii", "ignore").decode()
    n = re.sub(r"[^a-z ]", " ", n.lower())
    return " ".join(n.split())


class IdentityResolver:
    """Union-find over (email, login, full-name) keys for one repository."""

    def __init__(self, manual=None):
        self.parent = {}
        self.flags = {}
        self.manual = manual or {}  # email -> canonical id

    def _find(self, k):
        self.parent.setdefault(k, k)
        while self.parent[k] != k:
            self.parent[k] = self.parent[self.parent[k]]
            k = self.parent[k]
        return k

    def _union(self, a, b, flag=None):
        ra, rb = self._find(a), self._find(b)
        if ra != rb:
            self.parent[rb] = ra
            if flag:
                self.flags[ra] = flag

    def add(self, name, email, login=None):
        email = (email or "").strip().lower()
        key = "e:" + email if email else "n:" + norm_name(name)
        self._find(key)
        m = NOREPLY.match(email)
        if m:
            self._union("l:" + m.group(1).lower(), key)
        if login:
            self._union("l:" + login.lower(), key)
        nn = norm_name(name)
        if len(nn.split()) >= 2 and nn not in GENERIC_NAMES:
            self._union("n:" + nn, key, flag="name-match")
        if email in self.manual:
            self._union("m:" + self.manual[email], key, flag="manual")
        return key

    def canonical(self, name, email, login=None):
        return self._find(self.add(name, email, login))
