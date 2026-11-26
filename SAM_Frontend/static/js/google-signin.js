// "Sign in with Google": Google's own button (Google Identity Services).
// When it loads, it takes the place of SAM's Continue with Google form, which stays as the
// fallback when JavaScript or Google's script is blocked. The ID token Google hands back is
// posted to SAM, which checks its signature, audience, expiry and nonce before signing in.
(function () {
  var box = document.getElementById("google-button");
  if (!box) return;

  window.samGoogleButtonReady = function () {
    if (!window.google || !google.accounts || !google.accounts.id) return;
    var form = document.getElementById("google-button-form");
    google.accounts.id.initialize({
      client_id: box.dataset.clientId,
      nonce: box.dataset.nonce,
      callback: function (response) {
        form.elements.credential.value = response.credential;
        form.submit();
      },
    });
    var dark = window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches;
    google.accounts.id.renderButton(box, {
      type: "standard",
      theme: dark ? "filled_black" : "outline",
      size: "large",
      text: box.dataset.signup ? "signup_with" : "continue_with",
      shape: "rectangular",
      logo_alignment: "center",
      width: Math.min(400, Math.max(200, Math.round(box.parentElement.clientWidth))),
    });
    box.hidden = false;
    var fallback = document.getElementById("google-redirect-form");
    if (fallback) fallback.hidden = true;
  };
  if (window.google && google.accounts && google.accounts.id) window.samGoogleButtonReady();
})();
