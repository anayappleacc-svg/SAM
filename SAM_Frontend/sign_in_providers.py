"""sign_in_providers.py: "Continue with Google".

It uses OpenID Connect. SAM sends the student to Google, they sign in there,
and Google sends them back to SAM with a signed ID token that
says who they are. SAM checks the token's signature, audience, issuer, expiry
and one-time nonce before trusting it.

Setup (step by step in the README): GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET.

Until these are set, the button opens a clearly labelled demo sign-in instead,
so the rest of SAM can still be tried.
"""

import json
import os
import urllib.error
import urllib.parse
import urllib.request

import jwt

PROVIDERS = {
    "google": {
        "label": "Google",
        "authorize": "https://accounts.google.com/o/oauth2/v2/auth",
        "token": "https://oauth2.googleapis.com/token",
        "keys": "https://www.googleapis.com/oauth2/v3/certs",
        "issuers": ["https://accounts.google.com", "accounts.google.com"],
        "client_id": "GOOGLE_CLIENT_ID",
    },
}

_key_clients = {}


class SignInError(Exception):
    """Something went wrong signing in. The message is safe to show the student."""


def _client_id(provider):
    return os.environ.get(PROVIDERS[provider]["client_id"], "").strip()


def is_configured(provider):
    return bool(_client_id(provider) and os.environ.get("GOOGLE_CLIENT_SECRET", "").strip())


def start_url(provider, redirect_uri, state, nonce):
    """Where to send the student to sign in."""
    p = PROVIDERS[provider]
    params = {"client_id": _client_id(provider), "redirect_uri": redirect_uri, "state": state, "nonce": nonce,
              "response_type": "code", "scope": "openid email profile", "prompt": "select_account"}
    return p["authorize"] + "?" + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)


def finish_google(code, redirect_uri, nonce):
    """Swap Google's one-time code for the student's ID token, then check it."""
    body = urllib.parse.urlencode({
        "code": code,
        "client_id": _client_id("google"),
        "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET", "").strip(),
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    }).encode()
    req = urllib.request.Request(PROVIDERS["google"]["token"], data=body, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=15) as res:
            tokens = json.load(res)
    except (urllib.error.URLError, ValueError) as e:
        raise SignInError("Google sign-in didn’t finish. Try again.") from e
    return _identity("google", tokens.get("id_token", ""), nonce)


def _identity(provider, token, nonce):
    p = PROVIDERS[provider]
    label = p["label"]
    if not token:
        raise SignInError(f"{label} didn’t send back a sign-in. Try again.")
    try:
        if p["keys"] not in _key_clients:
            _key_clients[p["keys"]] = jwt.PyJWKClient(p["keys"], cache_keys=True, lifespan=3600)
        key = _key_clients[p["keys"]].get_signing_key_from_jwt(token)
        claims = jwt.decode(
            token, key.key, algorithms=["RS256"], audience=_client_id(provider), issuer=p["issuers"],
            leeway=60, options={"require": ["exp", "iat", "iss", "aud", "sub"]},
        )
    except jwt.PyJWTError as e:
        raise SignInError(f"SAM couldn’t confirm your {label} sign-in. Try again.") from e
    if claims.get("nonce") != nonce:
        raise SignInError(f"That {label} sign-in has expired. Try again.")
    email = claims.get("email") or ""
    if email and str(claims.get("email_verified", "true")).lower() != "true":
        email = ""  # Never trust an email address Google hasn't verified.
    return {"sub": str(claims["sub"]), "email": email, "name": claims.get("given_name") or claims.get("name") or ""}
