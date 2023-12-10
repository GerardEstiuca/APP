/**
* jquery_snow - jQuery Snow Effect Plugin
*
* Available under MIT licence
*
@version 1 (21. Jan 2012)
@author Ivan lazarevic
@requires jquery_snow
@see http://workshop.rs
*
@params minSize - min size of snowflake, 10 by default
@params maxSize - max size of snowflake, 20 by default
@params newOn - frequency in ms of appearing of new snowflake, 500 by default
@params flakecolor - color of snowflake, #FFFFFF by default
@example $.fn.snow({ maxSize: 200, newOn: 1000});
*/

(function($) {
    $.fn.snow = function(options) {
        var settings = $.extend({
            minSize: 10,    // Tamaño mínimo del copo de nieve
            maxSize: 20,    // Tamaño máximo del copo de nieve
            newOn: 200,     // Frecuencia en milisegundos para añadir un nuevo copo de nieve
            flakeColor: "#FFFFFF"  // Color del copo de nieve
        }, options);

        var $flake = $('<div id="flake" />').css({'position': 'absolute', 'top': '-50px'}).html('&#10052;');

        var documentHeight = $(document).height();
        var documentWidth = $(document).width();

        setInterval(function() {
            var startPositionLeft = Math.random() * documentWidth;
            var startOpacity = 0.5 + Math.random();
            var sizeFlake = settings.minSize + Math.random() * settings.maxSize;
            var endPositionTop = documentHeight - 40;
            var endPositionLeft = startPositionLeft - 100 + Math.random() * 200;
            var durationFall = documentHeight * 10 + Math.random() * 5000;

            $flake.clone().appendTo('body').css({
                left: startPositionLeft,
                opacity: startOpacity,
                'font-size': sizeFlake,
                color: settings.flakeColor
            }).animate({
                top: endPositionTop,
                left: endPositionLeft,
                opacity: 0.2
            }, durationFall, 'linear', function() {
                $(this).remove()
            });
        }, settings.newOn);
    };
})(jQuery);



  
