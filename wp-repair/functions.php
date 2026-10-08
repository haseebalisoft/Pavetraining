<?php
/**
 * PAVE Training child theme functions.
 *
 * Enqueues, block styles, pattern category, image sizes, font preload and SEO/schema.
 */
if (!defined('ABSPATH')) { exit; }

if (!function_exists('pave_is_home')) {
    function pave_is_home() { return is_front_page() || is_page('home'); }
}

/** Cache-bust theme assets from filemtime (forces new mobile menu CSS). */
function pave_asset_ver($relative) {
    $path = get_stylesheet_directory() . '/' . ltrim($relative, '/');
    return file_exists($path) ? (string) filemtime($path) : wp_get_theme()->get('Version');
}
add_action('wp_enqueue_scripts', function () {
    // Fonts are self-hosted via theme.json fontFace (Geist + Inter).
}, 5);
add_action('after_setup_theme', function () {
    add_theme_support('post-thumbnails');
    add_theme_support('responsive-embeds');
    add_editor_style('style.css');
    add_image_size('pave-hero', 1600, 900, true);
    add_image_size('pave-card', 800, 500, true);
});

add_action('wp_enqueue_scripts', function () {
    $uri = get_stylesheet_directory_uri();
    $dir = get_stylesheet_directory();
    $ver = wp_get_theme()->get('Version');
    wp_enqueue_style('pave-training', $uri . '/style.css', array(), pave_asset_ver('style.css'));
    if (file_exists($dir . '/assets/js/pave.js')) {
        wp_enqueue_script('pave-training', $uri . '/assets/js/pave.js', array(), pave_asset_ver('assets/js/pave.js'), array('in_footer' => true, 'strategy' => 'defer'));
    }
}, 20);

add_action('init', function () {
    register_block_style('core/button', array('name' => 'pave-outline', 'label' => 'PAVE Outline'));
    if (function_exists('register_block_pattern_category')) {
        register_block_pattern_category('pave', array('label' => 'PAVE Training'));
    }
    if (function_exists('register_block_pattern')) {
        $pave_patterns = array(
            'hero'              => 'PAVE Hero',
            'features'          => 'PAVE Features',
            'search'            => 'PAVE Course Search',
            'courses'           => 'PAVE Course Cards',
            'featured-courses'  => 'PAVE Featured Courses',
            'services-home'     => 'PAVE Services Home',
            'why-us'            => 'PAVE Why Us',
            'trust-strip'       => 'PAVE Trust Strip',
            'accreditations'    => 'PAVE Accreditations',
            'reviews'           => 'PAVE Google Reviews',
            'cta'               => 'PAVE CTA',
            'enquire'           => 'PAVE Enquire',
            'testimonials'      => 'PAVE Testimonials',
            'contact'           => 'PAVE Contact',
        );
        $pave_base = get_stylesheet_directory() . '/patterns-src/';
        foreach ($pave_patterns as $pave_slug => $pave_title) {
            $pave_file = $pave_base . $pave_slug . '.html';
            if (file_exists($pave_file)) {
                register_block_pattern('pave-training/' . $pave_slug, array(
                    'title'         => $pave_title,
                    'categories'    => array('pave'),
                    'inserter'      => true,
                    'content'       => file_get_contents($pave_file),
                ));
            }
        }
    }
});

/* Font preloads: pave_typography_preload() in inc/typography.php */

add_action('wp_head', function () {
    if (is_admin() || is_feed()) { return; }

    $site_name = 'PAVE Training';
    $home = home_url('/');
    $canonical = is_singular() ? get_permalink() : (is_post_type_archive('pave_course') ? get_post_type_archive_link('pave_course') : $home);

    // Prefer Rank Math stored values, then excerpt, then smart defaults.
    $title = wp_get_document_title();
    $desc = '';
    $focus = '';
    if (is_singular()) {
        $id = get_queried_object_id();
        $desc = (string) get_post_meta($id, 'rank_math_description', true);
        $custom_title = (string) get_post_meta($id, 'rank_math_title', true);
        $focus = (string) get_post_meta($id, 'rank_math_focus_keyword', true);
        if ($custom_title !== '') { $title = $custom_title; }
        if ($desc === '') {
            $excerpt = get_the_excerpt($id);
            $desc = $excerpt ? wp_strip_all_tags($excerpt) : '';
        }
        if ($desc === '') {
            $desc = wp_trim_words(wp_strip_all_tags(get_post_field('post_content', $id)), 32, '...');
        }
    } elseif (is_post_type_archive('pave_course')) {
        $title = 'Training Courses | PAVE Training';
        $desc = 'Browse accredited NPORS, EUSR, NVQ, CITB and health and safety courses from PAVE Training.';
    } else {
        $desc = 'Accredited construction, plant and industrial training. NPORS, EUSR, Streetworks, NVQ, CITB and health and safety courses nationwide.';
    }
    $desc = trim(preg_replace('/\s+/', ' ', wp_strip_all_tags($desc)));
    if ($desc === '') {
        $desc = 'Accredited construction, plant and industrial training from PAVE Training.';
    }

    // Unique SEO tags from Rank Math meta / excerpts (theme-owned for reliability).
    echo '<meta name="description" content="' . esc_attr($desc) . '">';
    echo '<link rel="canonical" href="' . esc_url($canonical) . '">';
    echo '<meta property="og:type" content="' . (is_singular('pave_course') ? 'article' : 'website') . '">';
    echo '<meta property="og:site_name" content="' . esc_attr($site_name) . '">';
    echo '<meta property="og:title" content="' . esc_attr($title) . '">';
    echo '<meta property="og:description" content="' . esc_attr($desc) . '">';
    echo '<meta property="og:url" content="' . esc_url($canonical) . '">';
    echo '<meta name="twitter:card" content="summary_large_image">';
    echo '<meta name="twitter:title" content="' . esc_attr($title) . '">';
    echo '<meta name="twitter:description" content="' . esc_attr($desc) . '">';
    if ($focus !== '') {
        echo '<meta name="keywords" content="' . esc_attr($focus) . '">';
    }

    $graph = array(
        array(
            '@type' => array('Organization', 'EducationalOrganization'),
            '@id' => $home . '#org',
            'name' => $site_name,
            'url' => $home,
            'description' => 'Accredited construction, plant and industrial training provider.',
            'areaServed' => 'GB',
            'slogan' => 'Industry Focused. Safety Led. Quality Driven.',
            'contactPoint' => array(
                '@type' => 'ContactPoint',
                'telephone' => '+44-1234-567890',
                'email' => 'info@pavetraining.co.uk',
                'contactType' => 'customer service',
            ),
        ),
        array(
            '@type' => 'WebSite',
            '@id' => $home . '#website',
            'url' => $home,
            'name' => $site_name,
            'publisher' => array('@id' => $home . '#org'),
            'potentialAction' => array(
                '@type' => 'SearchAction',
                'target' => $home . '?s={search_term_string}',
                'query-input' => 'required name=search_term_string',
            ),
        ),
    );

    if (is_singular('pave_course')) {
        $id = get_queried_object_id();
        $graph[] = array(
            '@type' => 'Course',
            '@id' => get_permalink($id) . '#course',
            'name' => get_the_title($id),
            'description' => $desc,
            'provider' => array('@id' => $home . '#org'),
            'url' => get_permalink($id),
        );
        $graph[] = array(
            '@type' => 'BreadcrumbList',
            'itemListElement' => array(
                array('@type' => 'ListItem', 'position' => 1, 'name' => 'Home', 'item' => $home),
                array('@type' => 'ListItem', 'position' => 2, 'name' => 'Courses', 'item' => get_post_type_archive_link('pave_course')),
                array('@type' => 'ListItem', 'position' => 3, 'name' => get_the_title($id), 'item' => get_permalink($id)),
            ),
        );
    } elseif (is_front_page()) {
        $graph[] = array(
            '@type' => 'BreadcrumbList',
            'itemListElement' => array(
                array('@type' => 'ListItem', 'position' => 1, 'name' => 'Home', 'item' => $home),
            ),
        );
    } elseif (is_singular('page')) {
        $graph[] = array(
            '@type' => 'BreadcrumbList',
            'itemListElement' => array(
                array('@type' => 'ListItem', 'position' => 1, 'name' => 'Home', 'item' => $home),
                array('@type' => 'ListItem', 'position' => 2, 'name' => get_the_title(), 'item' => get_permalink()),
            ),
        );
    }

    echo '<script type="application/ld+json">' . wp_json_encode(array('@context' => 'https://schema.org', '@graph' => $graph), JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE) . '</script>';
}, 5);

// Prefer Rank Math custom title in the document title when present.
add_filter('pre_get_document_title', function ($title) {
    if (!is_singular()) { return $title; }
    $custom = (string) get_post_meta(get_queried_object_id(), 'rank_math_title', true);
    return $custom !== '' ? $custom : $title;
}, 20);

add_action('init', function () {
    register_post_type('pave_course', array(
        'labels' => array(
            'name' => 'Courses',
            'singular_name' => 'Course',
            'add_new_item' => 'Add New Course',
            'edit_item' => 'Edit Course',
            'view_item' => 'View Course',
            'search_items' => 'Search Courses',
        ),
        'public' => true,
        'has_archive' => 'courses',
        'rewrite' => array('slug' => 'courses', 'with_front' => false),
        'menu_icon' => 'dashicons-welcome-learn-more',
        'supports' => array('title', 'editor', 'thumbnail', 'excerpt', 'revisions', 'custom-fields'),
        'show_in_rest' => true,
        'taxonomies' => array('pave_course_category'),
    ));

    register_taxonomy('pave_course_category', 'pave_course', array(
        'labels' => array(
            'name' => 'Course Categories',
            'singular_name' => 'Course Category',
        ),
        'public' => true,
        'hierarchical' => true,
        'rewrite' => array('slug' => 'course-category'),
        'show_in_rest' => true,
    ));
}, 5);

add_action('init', function () {
    $terms = array(
        'npors' => 'NPORS',
        'eusr' => 'EUSR',
        'nvq' => 'NVQ',
        'citb' => 'CITB',
        'streetworks' => 'Streetworks',
        'health-safety' => 'Health & Safety',
        'specialist' => 'Specialist',
    );
    foreach ($terms as $slug => $name) {
        if (!term_exists($slug, 'pave_course_category')) {
            wp_insert_term($name, 'pave_course_category', array('slug' => $slug));
        }
    }
}, 20);
add_action('init', function () {
    register_post_type('pave_enquiry', array(
        'labels' => array(
            'name' => 'Enquiries',
            'singular_name' => 'Enquiry',
            'menu_name' => 'Enquiries',
        ),
        'public' => false,
        'show_ui' => true,
        'show_in_menu' => true,
        'menu_icon' => 'dashicons-email-alt',
        'supports' => array('title', 'editor', 'custom-fields'),
        'capability_type' => 'post',
    ));
}, 6);

add_action('fluentform/submission_inserted', function ($entryId, $formData, $form) {
    $type = isset($formData['enquiry_type']) ? sanitize_text_field($formData['enquiry_type']) : 'general';
    $name = '';
    if (isset($formData['names']) && is_array($formData['names'])) {
        $name = trim(($formData['names']['first_name'] ?? '') . ' ' . ($formData['names']['last_name'] ?? ''));
    } elseif (!empty($formData['names'])) {
        $name = is_string($formData['names']) ? $formData['names'] : '';
    }
    $email = isset($formData['email']) ? sanitize_email($formData['email']) : '';
    $phone = isset($formData['phone']) ? sanitize_text_field($formData['phone']) : '';
    $message = isset($formData['message']) ? sanitize_textarea_field($formData['message']) : '';

    $post_id = wp_insert_post(array(
        'post_type' => 'pave_enquiry',
        'post_status' => 'publish',
        'post_title' => sprintf('%s — %s', strtoupper($type), $name ?: $email ?: 'Website enquiry'),
        'post_content' => $message,
    ));
    if ($post_id && !is_wp_error($post_id)) {
        update_post_meta($post_id, 'enquiry_type', $type);
        update_post_meta($post_id, 'enquiry_email', $email);
        update_post_meta($post_id, 'enquiry_phone', $phone);
        update_post_meta($post_id, 'enquiry_name', $name);
        update_post_meta($post_id, 'enquiry_status', 'New');
        update_post_meta($post_id, 'fluentform_entry_id', (int) $entryId);
    }

    $webhook = get_option('pave_power_automate_webhook');
    if ($webhook) {
        wp_remote_post($webhook, array(
            'timeout' => 12,
            'headers' => array('Content-Type' => 'application/json'),
            'body' => wp_json_encode(array(
                'source' => 'pavetraining.co.uk',
                'entry_id' => $entryId,
                'enquiry_type' => $type,
                'name' => $name,
                'email' => $email,
                'phone' => $phone,
                'message' => $message,
                'status' => 'New',
                'submitted_at' => current_time('c'),
            )),
        ));
    }
}, 20, 3);

add_action('rest_api_init', function () {
    register_rest_route('pave/v1', '/power-automate-webhook', array(
        'methods' => 'POST',
        'permission_callback' => function () {
            return current_user_can('manage_options');
        },
        'callback' => function ($request) {
            $url = esc_url_raw($request->get_param('url'));
            if (!$url) {
                return new WP_Error('missing_url', 'Provide url', array('status' => 400));
            }
            update_option('pave_power_automate_webhook', $url, false);
            return array('saved' => true, 'url' => $url);
        },
    ));
});
/** pave_security_performance */
add_action('init', function () {
    // Remove generator fingerprint
    remove_action('wp_head', 'wp_generator');
    add_filter('the_generator', '__return_empty_string');

    // Disable XML-RPC (common attack surface)
    add_filter('xmlrpc_enabled', '__return_false');

    // Disable author enumeration via ?author=1 redirects for guests
    if (!is_admin() && isset($_GET['author']) && !is_user_logged_in()) {
        wp_safe_redirect(home_url('/'), 301);
        exit;
    }
}, 1);

add_filter('rest_authentication_errors', function ($result) {
    return $result;
});

// Disallow file edit in admin for safer WP
if (!defined('DISALLOW_FILE_EDIT')) {
    // Cannot define late if already loaded; set via filter alternatives
}
add_filter('file_mod_allowed', function ($allowed, $context) {
    if ($context === 'capability_edit_themes' || $context === 'capability_edit_plugins') {
        return false;
    }
    return $allowed;
}, 10, 2);

add_action('send_headers', function () {
    if (is_admin()) { return; }
    header('X-Content-Type-Options: nosniff');
    header('X-Frame-Options: SAMEORIGIN');
    header('Referrer-Policy: strict-origin-when-cross-origin');
    header('Permissions-Policy: geolocation=(), microphone=(), camera=()');
    header('Cross-Origin-Opener-Policy: same-origin-allow-popups');
});

// Performance: dequeue unused front-end assets for guests
add_action('wp_enqueue_scripts', function () {
    if (is_admin()) { return; }
    wp_dequeue_style('wp-block-library-theme');
    wp_deregister_script('wp-embed');

    // Emoji scripts unused on marketing pages
    remove_action('wp_head', 'print_emoji_detection_script', 7);
    remove_action('wp_print_styles', 'print_emoji_styles');
    remove_action('admin_print_scripts', 'print_emoji_detection_script');
    remove_action('admin_print_styles', 'print_emoji_styles');
}, 100);

// Lazy-load attributes for content images (WP core supports; force decoding async)
add_filter('wp_get_attachment_image_attributes', function ($attr) {
    $attr['loading'] = $attr['loading'] ?? 'lazy';
    $attr['decoding'] = 'async';
    return $attr;
}, 20);

add_filter('wp_lazy_loading_enabled', '__return_true');

// Preload hero image on front page only
add_action('wp_head', function () {
    if (!is_front_page()) { return; }
    $hero = home_url('/wp-content/uploads/customer-hero-paver.png');
    echo '<link rel="preload" as="image" href="' . esc_url($hero) . '" fetchpriority="high">';
}, 2);

// Hide login errors precision
add_filter('login_errors', function () {
    return 'Invalid login details.';
});
/** Portal-matching homepage offer slider */
add_shortcode('pave_offer_slider', 'pave_offer_slider_shortcode');
function pave_offer_slider_shortcode() {
    $slides = array(
        array(
            'tag' => 'PAVE Training',
            'title' => 'Paving the way in industry',
            'desc' => 'High quality, accredited training courses for the construction, plant and industrial sectors.',
            'cta' => 'View courses',
            'href' => '#courses',
            'bg' => home_url('/wp-content/uploads/customer-hero-paver.png'),
        ),
        array(
            'tag' => 'Accreditations',
            'title' => 'Industry accredited training',
            'desc' => 'NOCN · NPORS · EUSR Approved Trainer · Street Works Qualifications Register — trusted credentials for your workforce.',
            'cta' => 'Explore NPORS',
            'href' => '/npors/',
            'bg' => home_url('/wp-content/uploads/customer-hero-accreditations.png'),
        ),
        array(
            'tag' => 'Limited time',
            'title' => '15% off NPORS refresher courses',
            'desc' => 'Book any NPORS refresher for your team and save 15% — ask our team for current dates and eligibility.',
            'cta' => 'Enquire now',
            'href' => '/contact/?enquiry=npors',
            'bg' => home_url('/wp-content/uploads/customer-hero-discount.png'),
        ),
    );

    ob_start();
    $first_bg = esc_url($slides[0]['bg']);
    echo '<div class="pave-offer-slider" data-pave-slider tabindex="0" role="region" aria-roledescription="carousel" aria-label="PAVE Training highlights" style="--pave-slide-bg:url(\'' . $first_bg . '\')">';
    echo '<div class="pave-offer-slider__media" aria-hidden="true"></div>';
    echo '<div class="pave-offer-slider__shade" aria-hidden="true"></div>';
    foreach ($slides as $i => $slide) {
        $active = $i === 0 ? ' is-active' : '';
        echo '<div class="pave-offer-slider__slide' . $active . '" data-slide data-bg="' . esc_url($slide['bg']) . '" aria-hidden="' . ($i === 0 ? 'false' : 'true') . '">';
        echo '<div class="pave-offer-slider__content">';
        echo '<span class="pave-offer-slider__eyebrow">' . esc_html($slide['tag']) . '</span>';
        echo '<h1 class="pave-offer-slider__title">' . esc_html($slide['title']) . '</h1>';
        echo '<p class="pave-offer-slider__desc">' . esc_html($slide['desc']) . '</p>';
        echo '<a class="pave-offer-slider__cta" href="' . esc_url($slide['href']) . '"><span>' . esc_html($slide['cta']) . '</span><span class="pave-offer-slider__cta-arrow" aria-hidden="true">→</span></a>';
        echo '</div></div>';
    }
    echo '<button type="button" class="pave-offer-slider__arrow pave-offer-slider__arrow--prev" data-prev aria-label="Previous slide"><span aria-hidden="true">‹</span></button>';
    echo '<button type="button" class="pave-offer-slider__arrow pave-offer-slider__arrow--next" data-next aria-label="Next slide"><span aria-hidden="true">›</span></button>';
    echo '<div class="pave-offer-slider__dots" aria-label="Choose slide">';
    foreach ($slides as $i => $slide) {
        $active = $i === 0 ? ' is-active' : '';
        echo '<button type="button" class="pave-offer-slider__dot' . $active . '" data-dot aria-label="Show slide ' . ($i + 1) . ': ' . esc_attr($slide['title']) . '"' . ($i === 0 ? ' aria-current="true"' : '') . '></button>';
    }
    echo '</div></div>';
    return ob_get_clean();
}
add_filter('render_block', 'pave_render_hero_slider_block', 20, 2);
function pave_render_hero_slider_block($block_content, $block) {
    if (($block['blockName'] ?? '') !== 'core/group') {
        return $block_content;
    }
    $class = (string) ($block['attrs']['className'] ?? '');
    if (strpos($class, 'pave-hero-wrap') === false) {
        return $block_content;
    }
    if (function_exists('pave_hero_slider_shortcode')) {
        $slider = pave_hero_slider_shortcode();
    } elseif (function_exists('pave_offer_slider_shortcode')) {
        $slider = pave_offer_slider_shortcode();
    } else {
        return $block_content;
    }
    // Keep outer section classes/styles from block render, replace inner
    if (preg_match('/^(\s*<section\b[^>]*>)/i', $block_content, $m)) {
        return $m[1] . $slider . '</section>';
    }
    return '<section class="wp-block-group alignfull pave-hero-wrap">' . $slider . '</section>';
}
add_shortcode('pave_accreditations_panel', 'pave_accreditations_panel_shortcode');
function pave_accreditations_panel_shortcode() {
    $bg = esc_url(home_url('/wp-content/uploads/customer-hero-accreditations.png'));
    $marks = array(
        array('NOCN', 'Awarding body'),
        array('NPORS', 'Plant ops'),
        array('EUSR', 'Utilities'),
        array('SWQR', 'Street works'),
        array('CITB', 'Construction'),
        array('NVQ', 'Vocational'),
    );
    ob_start();
    ?>
    <div class="pave-acred-panel">
      <div class="pave-acred-panel__copy">
        <span class="pave-acred-panel__tag">Trusted credentials</span>
        <h2 class="pave-acred-panel__title">Accreditations and memberships</h2>
        <p class="pave-acred-panel__desc">Recognised training pathways across plant, utilities, street works and vocational qualifications — built for contractors who need audit-ready competence.</p>
        <ul class="pave-acred-panel__marks" aria-label="Accreditation marks">
          <?php foreach ($marks as $m) : ?>
            <li class="pave-acred-panel__mark">
              <span class="pave-acred-panel__mark-code"><?php echo esc_html($m[0]); ?></span>
              <span class="pave-acred-panel__mark-label"><?php echo esc_html($m[1]); ?></span>
            </li>
          <?php endforeach; ?>
        </ul>
        <div class="pave-acred-panel__actions">
          <a class="pave-acred-panel__cta" href="/courses/">Browse courses <span aria-hidden="true">→</span></a>
          <a class="pave-acred-panel__link" href="/about/">Why PAVE</a>
        </div>
      </div>
      <figure class="pave-acred-panel__visual">
        <div class="pave-acred-panel__frame" style="--pave-acred-bg:url('<?php echo $bg; ?>')" role="img" aria-label="Heavy plant operator delivering accredited PAVE training"></div>
        <figcaption class="pave-acred-panel__chip">NOCN · NPORS · EUSR · Street Works</figcaption>
      </figure>
    </div>
    <?php
    return ob_get_clean();
}
add_filter('render_block', 'pave_render_accreditations_block', 21, 2);
function pave_render_accreditations_block($block_content, $block) {
    if (($block['blockName'] ?? '') !== 'core/group') {
        return $block_content;
    }
    $class = (string) ($block['attrs']['className'] ?? '');
    if (strpos($class, 'pave-accreditations') === false || strpos($class, 'pave-hero-wrap') !== false) {
        return $block_content;
    }
    if (!function_exists('pave_accreditations_panel_shortcode')) {
        return $block_content;
    }
    $panel = pave_accreditations_panel_shortcode();
    if (preg_match('/^(\s*<section\b[^>]*>)/i', $block_content, $m)) {
        return $m[1] . $panel . '</section>';
    }
    return '<section class="wp-block-group alignfull pave-accreditations">' . $panel . '</section>';
}
/** Prefer course results for site search. */
add_action('pre_get_posts', 'pave_search_courses');
function pave_search_courses($query) {
    if (is_admin() || !$query->is_main_query() || !$query->is_search()) {
        return;
    }
    $ptype = isset($_GET['post_type']) ? sanitize_key(wp_unslash($_GET['post_type'])) : '';
    if ($ptype === 'pave_course') {
        $query->set('post_type', 'pave_course');
        return;
    }
    // Default site search: pages + courses
    $query->set('post_type', array('post', 'page', 'pave_course'));
}

add_shortcode('pave_trust_strip', 'pave_trust_strip_shortcode');
function pave_trust_strip_shortcode() {
    $items = array(
        array('label' => 'Delivery', 'value' => '500+ Courses Delivered', 'icon' => 'courses'),
        array('label' => 'Coverage', 'value' => 'Nationwide Training Centres', 'icon' => 'pin'),
        array('label' => 'Funding', 'value' => 'CITB Grant Eligible', 'icon' => 'grant'),
        array('label' => 'Standards', 'value' => 'NPORS Accredited', 'icon' => 'badge'),
    );
    $icons = array(
        'courses' => '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M4 6.5h16M4 12h16M4 17.5h10" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
        'pin' => '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 21s7-5.4 7-11a7 7 0 1 0-14 0c0 5.6 7 11 7 11z" stroke="currentColor" stroke-width="1.8"/><circle cx="12" cy="10" r="2.2" stroke="currentColor" stroke-width="1.8"/></svg>',
        'grant' => '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 3v18M7 8.5c0-1.9 2.2-3 5-3s5 1.1 5 3-2.2 3-5 3-5 1.1-5 3 2.2 3 5 3 5-1.1 5-3" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
        'badge' => '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 3l2.4 4.9 5.4.8-3.9 3.8.9 5.4L12 15.9 7.2 18l.9-5.4L4.2 8.7l5.4-.8L12 3z" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>',
    );
    ob_start();
    echo '<div class="pave-trust-strip" aria-label="PAVE Training credentials">';
    echo '<div class="pave-trust-strip__inner">';
    foreach ($items as $item) {
        echo '<div class="pave-trust-strip__item">';
        echo '<span class="pave-trust-strip__icon">' . ($icons[$item['icon']] ?? '') . '</span>';
        echo '<span class="pave-trust-strip__copy">';
        echo '<span class="pave-trust-strip__label">' . esc_html($item['label']) . '</span>';
        echo '<span class="pave-trust-strip__value">' . esc_html($item['value']) . '</span>';
        echo '</span></div>';
    }
    echo '</div></div>';
    return ob_get_clean();
}

add_filter('render_block', 'pave_render_trust_strip_block', 22, 2);
function pave_render_trust_strip_block($block_content, $block) {
    if (($block['blockName'] ?? '') !== 'core/group') {
        return $block_content;
    }
    $class = (string) ($block['attrs']['className'] ?? '');
    if (strpos($class, 'pave-trust-strip-wrap') === false) {
        return $block_content;
    }
    $strip = pave_trust_strip_shortcode();
    if (preg_match('/^(\s*<section\b[^>]*>)/i', $block_content, $m)) {
        return $m[1] . $strip . '</section>';
    }
    return '<section class="wp-block-group alignfull pave-trust-strip-wrap">' . $strip . '</section>';
}

if (file_exists(__DIR__ . '/inc/course-system.php')) {
    require_once __DIR__ . '/inc/course-system.php';
}
if (file_exists(__DIR__ . '/inc/typography.php')) {
    require_once __DIR__ . '/inc/typography.php';
}
if (file_exists(__DIR__ . '/inc/site-content.php')) {
    require_once __DIR__ . '/inc/site-content.php';
}
if (file_exists(__DIR__ . '/inc/static-pages.php')) {
    require_once __DIR__ . '/inc/static-pages.php';
}
if (file_exists(__DIR__ . '/inc/home-sections.php')) {
    require_once __DIR__ . '/inc/home-sections.php';
}
if (file_exists(__DIR__ . '/inc/home-cats-popular.php')) {
    require_once __DIR__ . '/inc/home-cats-popular.php';
}

/**
 * New schema-driven backend. Inert until the option is set to '1'.
 * Enable:  update_option('pave_new_backend', '1');
 * Disable: update_option('pave_new_backend', '0');
 */
if (get_option('pave_new_backend') === '1') {
    $pave_new = array(
        'course-schema',
        'course-helpers',
        'site-options',
        'course-categories',
        'section-library',
        'course-components',
        'premium-pass',
        'course-images',
        'mega-nav',
        'course-admin-ui',
        'hero-slider',
        'hero-admin',
    );
    foreach ($pave_new as $pave_file) {
        $pave_path = get_stylesheet_directory() . '/inc/' . $pave_file . '.php';
        if (file_exists($pave_path)) {
            require_once $pave_path;
        }
    }
    unset($pave_new, $pave_file, $pave_path);
}

add_filter('body_class', function ($classes) {
    $classes[] = 'pave-fonts-ready';
    return $classes;
});

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
