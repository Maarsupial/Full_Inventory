var coll = document.getElementsByClassName("collapsible");
var i;

for (i = 0; i < coll.length; i++) {
    coll[i].addEventListener("click", function() {
        this.classList.toggle("active");
        var content = this.nextElementSibling;
        var current = this.nextElementSibling.id;

        var liCount = document.querySelectorAll('#' + current + ' li').length;
        var brCount = document.querySelectorAll('#' + current + ' br').length;

        var lineCount = liCount + brCount;
        lineCount *= 19;
        lineCount += 32;
        
        var height = document.getElementById(current).clientHeight;

        if(height == 0) {
            content.style.height = lineCount + "px";
        } else {
            content.style.height = "0px";
        }
    });
}
