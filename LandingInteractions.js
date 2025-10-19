(function (factory) {
  if (typeof window !== 'undefined' && window.jQuery) {
    factory(window.jQuery);
  } else {
    if (typeof console !== 'undefined' && console.warn) {
      console.warn('LandingInteractions requires jQuery to run. Animations are disabled.');
    }
  }
})(function ($) {
  function applyRevealAnimations() {
    var revealSelector =
      '.metric-card, .card, .mission-card, .timeline-step, .values-grid .card, .workflow-card, ' +
      '.detail-card, .module-card, .story-card, .persona-card, .playbook-card, .stack-card, ' +
      '.panel-copy, .panel-media, .phone-mockups img, .hero-floating-card';

    var $revealables = $(revealSelector);

    $revealables.each(function () {
      var $element = $(this);
      if (!$element.hasClass('will-reveal')) {
        $element.addClass('will-reveal');
      }
    });

    function revealOnScroll() {
      var viewportBottom = $(window).scrollTop() + $(window).height() * 0.88;
      $revealables.each(function () {
        var $element = $(this);
        if ($element.hasClass('is-visible')) {
          return;
        }

        var elementTop = $element.offset().top;
        if (elementTop <= viewportBottom) {
          $element.addClass('is-visible');
        }
      });
    }

    $(window).on('scroll.reveal resize.reveal', revealOnScroll);
    revealOnScroll();
  }

  function applyMetricCounters() {
    $('[data-counter]').each(function (index) {
      var $counter = $(this);
      var targetValue = parseFloat($counter.data('counter'));

      if (isNaN(targetValue)) {
        return;
      }

      var prefix = $counter.data('prefix') || '';
      var suffix = $counter.data('suffix') || '';
      var decimals = parseInt($counter.data('decimals'), 10);
      var duration = parseInt($counter.data('duration'), 10);
      var startValue = parseFloat($counter.data('start'));

      if (isNaN(decimals)) {
        decimals = 0;
      }

      if (isNaN(duration)) {
        duration = 1600;
      }

      if (isNaN(startValue)) {
        startValue = 0;
      }

      var hasAnimated = false;
      var namespace = '.counterWatch' + index;

      function formatValue(value) {
        if (decimals === 0) {
          return Math.round(value).toString();
        }

        return value.toFixed(decimals);
      }

      function animateCounter() {
        $({ value: startValue }).animate(
          { value: targetValue },
          {
            duration: duration,
            easing: 'swing',
            step: function (now) {
              $counter.text(prefix + formatValue(now) + suffix);
            },
            complete: function () {
              $counter.text(prefix + formatValue(targetValue) + suffix);
            },
          }
        );
      }

      function checkVisibility() {
        if (hasAnimated) {
          return;
        }

        var viewportBottom = $(window).scrollTop() + $(window).height() * 0.9;
        if ($counter.offset().top <= viewportBottom) {
          hasAnimated = true;
          $(window).off('scroll' + namespace, checkVisibility);
          $(window).off('resize' + namespace, checkVisibility);
          animateCounter();
        }
      }

      $(window).on('scroll' + namespace, checkVisibility);
      $(window).on('resize' + namespace, checkVisibility);
      checkVisibility();
    });
  }

  $(function () {
    applyRevealAnimations();
    applyMetricCounters();
  });
});
