<?php
/* PAVE COMING SOON — public locked; admins + secret preview unlock */
if (!defined('PAVE_PREVIEW_SECRET')) {
    define('PAVE_PREVIEW_SECRET', 'pave-preview-af39212fb2b53178');
}
add_action('template_redirect', function () {
    // Force public Coming Soon on (ignore option until go-live).
    if (is_admin()) {
        return;
    }
    $uri = $_SERVER['REQUEST_URI'] ?? '';
    if (preg_match('#wp-login\.php|wp-admin|wp-cron\.php|admin-ajax\.php|wp-json#i', $uri)) {
        return;
    }

    $cookie = 'pave_site_preview';
    if (isset($_GET['preview']) && defined('PAVE_PREVIEW_SECRET') && hash_equals(PAVE_PREVIEW_SECRET, (string) $_GET['preview'])) {
        $token = hash_hmac('sha256', PAVE_PREVIEW_SECRET, AUTH_SALT);
        setcookie($cookie, $token, array(
            'expires' => time() + 14 * DAY_IN_SECONDS,
            'path' => defined('COOKIEPATH') && COOKIEPATH ? COOKIEPATH : '/',
            'domain' => defined('COOKIE_DOMAIN') ? COOKIE_DOMAIN : '',
            'secure' => is_ssl(),
            'httponly' => true,
            'samesite' => 'Lax',
        ));
        $_COOKIE[$cookie] = $token;
        wp_safe_redirect(home_url('/?previewed=1'));
        exit;
    }

    $unlocked = false;
    if (!empty($_COOKIE[$cookie]) && defined('PAVE_PREVIEW_SECRET')) {
        $expected = hash_hmac('sha256', PAVE_PREVIEW_SECRET, AUTH_SALT);
        $unlocked = hash_equals($expected, (string) $_COOKIE[$cookie]);
    }
    if ($unlocked) {
        return;
    }
    if (is_user_logged_in() && current_user_can('edit_posts')) {
        return;
    }

    status_header(200);
    nocache_headers();
    $logo = function_exists('pave_site') ? pave_site('logo_url') : content_url('/uploads/pave-logo.png');
    $phone = function_exists('pave_site') ? pave_site('phone') : '+44 7305 901428';
    $email = function_exists('pave_site') ? pave_site('email') : 'info@pavetraining.co.uk';
    $cs_title = function_exists('pave_site') ? pave_site('coming_soon_title') : 'Coming soon';
    $cs_text = function_exists('pave_site') ? pave_site('coming_soon_text') : "We're preparing the new site. Accredited plant, utilities and vocational training — launching shortly.";
    $tel = function_exists('pave_site') ? pave_site('phone_tel') : '+447305901428';
    ?><!DOCTYPE html>
<html <?php language_attributes(); ?>>
<head>
<meta charset="<?php bloginfo('charset'); ?>">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Coming Soon — PAVE Training</title>
<style>
:root{--pave-green:#81CF43;--pave-dark:#0d0d0d}
*{box-sizing:border-box}
html,body{margin:0;min-height:100%;background:var(--pave-dark);color:#fff;font-family:Inter,system-ui,sans-serif}
body{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:2rem 1.25rem;background:radial-gradient(70% 55% at 50% 0%,rgba(129,207,67,.16),transparent 60%),#0d0d0d}
.cs{width:min(34rem,100%);text-align:center}
.cs__logo{display:inline-flex;align-items:center;justify-content:center;width:5.5rem;height:5.5rem;padding:.55rem;border-radius:14px;background:#fff;margin:0 auto 1.35rem}
.cs__logo img{display:block;width:100%;height:auto;object-fit:contain}
.cs__eyebrow{margin:0 0 .55rem;color:var(--pave-green);font-size:.72rem;font-weight:700;letter-spacing:.16em;text-transform:uppercase}
h1{margin:0;font-size:clamp(2rem,6vw,2.75rem);font-weight:700;letter-spacing:-.03em;line-height:1.08}
p{margin:1rem auto 0;max-width:36ch;color:rgba(255,255,255,.72);font-size:1.05rem;line-height:1.55}
.cs__contact{margin-top:1.75rem;display:flex;flex-wrap:wrap;gap:.75rem;justify-content:center}
.cs__contact a{display:inline-flex;align-items:center;justify-content:center;min-height:2.65rem;padding:.65rem 1.1rem;border-radius:999px;text-decoration:none;font-size:.9rem;font-weight:700}
.cs__contact a.primary{background:var(--pave-green);color:#111}
.cs__contact a.ghost{border:1px solid rgba(255,255,255,.28);color:#fff}
</style>
</head>
<body>
<main class="cs">
  <div class="cs__logo"><img src="<?php echo esc_url($logo); ?>" alt="PAVE Training" width="88" height="88"></div>
  <p class="cs__eyebrow">PAVE Training</p>
  <h1><?php echo esc_html($cs_title); ?></h1>
  <p><?php echo esc_html($cs_text); ?></p>
  <div class="cs__contact">
    <a class="primary" href="tel:<?php echo esc_attr($tel); ?>"><?php echo esc_html($phone); ?></a>
    <a class="ghost" href="mailto:<?php echo esc_attr($email); ?>"><?php echo esc_html($email); ?></a>
  </div>
</main>
</body>
</html><?php
    exit;
}, 0);
