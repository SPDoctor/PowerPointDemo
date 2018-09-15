// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

function doSomething(event) {
    Office.context.document.setSelectedDataAsync("Did something!");
    event.completed();
}