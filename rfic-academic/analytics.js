/* Academic website analytics. Public measurement identifier only. */
(() => {
  'use strict';
  const measurementId = 'G-92S1ZHRRC6';
  const siteTrack = 'RFIC';
  const basePath = '/rfic-academic/';
  const ownerKey = 'academic_analytics_owner';
  const consentKey = 'academic_analytics_consent_' + siteTrack;
  const read = key => { try { return localStorage.getItem(key); } catch (_) { return null; } };
  const write = (key, value) => { try { localStorage.setItem(key, value); } catch (_) {} };

  // The URL fragment is browser-local and never part of an HTTP request.
  if (location.hash === '#analytics-owner-on' || location.hash === '#analytics-owner-off') {
    write(ownerKey, location.hash.endsWith('-on') ? '1' : '0');
    history.replaceState(null, '', location.pathname + location.search);
  }
  window.addEventListener('hashchange', () => {
    if (location.hash === '#analytics-owner-on' || location.hash === '#analytics-owner-off') {
      write(ownerKey, location.hash.endsWith('-on') ? '1' : '0');
      history.replaceState(null, '', location.pathname + location.search);
      location.reload();
    }
  });

  let started = false;
  const isOwner = () => read(ownerKey) === '1';
  const analyticsConsent = () => read(consentKey) === 'granted' ? 'granted' : 'denied';

  function cleanLocation() {
    const url = new URL(location.origin + location.pathname);
    const allowed = {
      utm_source:['gmail'],
      utm_medium:['email'],
      utm_campaign:['fall2027_professor_outreach'],
      utm_content:['rfic_signature','eda_signature','rfic_body','eda_body']
    };
    const actual = new URLSearchParams(location.search);
    for (const [key, values] of Object.entries(allowed)) {
      if (values.includes(actual.get(key))) url.searchParams.set(key, actual.get(key));
    }
    return url.href;
  }

  function start() {
    if (started || isOwner() || location.hostname !== 'aryanaditta.github.io') return;
    started = true;
    window.dataLayer = window.dataLayer || [];
    window.gtag = window.gtag || function () { window.dataLayer.push(arguments); };
    const gtag = window.gtag;

    // Privacy-preserving default: no Analytics storage/cookies unless explicitly enabled.
    gtag('consent', 'default', {
      analytics_storage: analyticsConsent(),
      ad_storage: 'denied',
      ad_user_data: 'denied',
      ad_personalization: 'denied'
    });
    gtag('js', new Date());
    gtag('set', {
      site_track: siteTrack,
      allow_google_signals: false,
      allow_ad_personalization_signals: false
    });
    gtag('config', measurementId, {
      site_track: siteTrack,
      cookie_prefix: 'academic_' + siteTrack.toLowerCase(),
      cookie_path: basePath,
      cookie_flags: 'SameSite=Lax;Secure',
      page_location: cleanLocation(),
      page_referrer: document.referrer ? new URL(document.referrer).origin : '',
      allow_google_signals: false,
      allow_ad_personalization_signals: false
    });

    const script = document.createElement('script');
    script.async = true;
    script.src = 'https://www.googletagmanager.com/gtag/js?id=' + measurementId;
    document.head.appendChild(script);
  }

  function event(name, params) {
    if (!started || isOwner()) return;
    window.gtag('event', name, Object.assign({
      send_to: measurementId,
      site_track: siteTrack,
      page_path: location.pathname,
      transport_type: 'beacon'
    }, params));
  }

  document.addEventListener('click', e => {
    const a = e.target.closest && e.target.closest('a[href]');
    if (!a) return;
    const url = new URL(a.href, location.href);
    if (!/^https?:$/.test(url.protocol)) return;
    const path = url.pathname;

    if (url.origin === location.origin && path.startsWith(basePath) && /_CV\.pdf$/i.test(path)) {
      event('cv_click', {cv_path:path});
    } else if (url.origin === location.origin && path === basePath + 'research.html') {
      event('research_click', {});
    } else if (url.origin === location.origin && path === basePath + 'publications.html') {
      event('publications_click', {});
    } else if (/overview.*\.(png|jpe?g|webp|svg)$/i.test(path)) {
      event('research_overview_click', {overview_path:path});
    } else {
      const destination = url.hostname === 'scholar.google.com' ? 'scholar' : /(^|\.)linkedin\.com$/.test(url.hostname) ? 'linkedin' : url.hostname === 'github.com' ? 'github' : null;
      if (destination) event('external_profile_click', {destination});
    }
  });

  function clearAnalyticsCookies() {
    for (const item of document.cookie.split(';')) {
      const name = item.split('=')[0].trim();
      if (name.startsWith('academic_' + siteTrack.toLowerCase())) {
        for (const domain of ['', '; domain=' + location.hostname, '; domain=.' + location.hostname]) {
          document.cookie = name + '=; max-age=0; path=' + basePath + domain;
        }
      }
    }
  }

  // Changes made in another tab should take effect immediately.
  window.addEventListener('storage', e => {
    if (e.key === ownerKey || e.key === consentKey) location.reload();
  });

  // Analytics choices are available only on request from the footer; no first-visit popup.
  function choice() {
    if (isOwner()) return;
    const previous = document.getElementById('analytics-choice');
    if (previous) previous.remove();

    const panel = document.createElement('aside');
    panel.id = 'analytics-choice';
    panel.setAttribute('aria-label','Analytics choices');
    panel.style.cssText = 'position:fixed;bottom:12px;left:12px;right:12px;max-width:620px;margin:auto;padding:14px 16px;background:#fff;color:#203746;border:1px solid #b9cbd3;border-radius:8px;box-shadow:0 3px 18px #0002;z-index:1000;font:14px/1.5 system-ui';
    panel.innerHTML = '<span>Analytics runs without Analytics cookies by default. You may optionally allow Analytics cookies for more accurate session and returning-visitor measurement. </span><a href="privacy.html">Privacy</a><div style="display:flex;gap:12px;margin-top:8px"><button type="button" data-choice="granted">Allow analytics cookies</button><button type="button" data-choice="denied">Keep cookieless analytics</button></div>';

    panel.addEventListener('click', e => {
      const button = e.target.closest('button[data-choice]');
      if (!button) return;
      const consent = button.dataset.choice;
      write(consentKey, consent);
      if (window.gtag) {
        window.gtag('consent', 'update', {
          analytics_storage: consent,
          ad_storage: 'denied',
          ad_user_data: 'denied',
          ad_personalization: 'denied'
        });
      }
      if (consent === 'denied') clearAnalyticsCookies();
      panel.remove();
    });
    document.body.appendChild(panel);
  }

  const settings = document.querySelector('[data-analytics-settings]');
  if (settings) settings.addEventListener('click', e => {
    e.preventDefault();
    choice();
  });

  start();
})();
