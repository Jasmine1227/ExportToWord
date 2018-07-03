if (typeof jQuery !== "undefined" && typeof saveAs !== "undefined") {
    (function ($) {
        $.fn.wordExport = function (fileName) {
            fileName = typeof fileName !== 'undefined' ? fileName : "jQuery-Word-Export";
            var static = {
                mhtml: {
                    top: "Mime-Version: 1.0\nContent-Base: " + location.href + "\nContent-Type: Multipart/related; boundary=\"NEXT.ITEM-BOUNDARY\";type=\"text/html\"\n\n--NEXT.ITEM-BOUNDARY\nContent-Type: text/html; charset=\"utf-8\"\nContent-Location: " + location.href + "\n\n<!DOCTYPE html>\n<html>\n_html_</html>",
                    head: "<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n<style>\n_styles_\n</style>\n</head>\n",
                    body: "<body>_body_</body>"
                }
            };
            var options = {
                maxWidth: 624
            };
            // Clone selected element before manipulating it  
            var markup = $(this).clone();

            // Remove hidden elements from the output  
            markup.each(function () {
                var self = $(this);
                if (self.is(':hidden'))
                    self.remove();
            });

            // Embed all images using Data URLs  
            var images = Array();
            var img = markup.find('img');
            for (var i = 0; i < img.length; i++) {
                // Calculate dimensions of output image  
                var w = Math.min(img[i].width, options.maxWidth);
                var h = img[i].height * (w / img[i].width);
                var img_id = "#" + img[i].id;
                $('<canvas>').attr("id", "test_word_img_" + i).width(w).height(h).insertAfter(img_id);
                // Create canvas for converting image to data URL  
                //var canvas = document.createElement("CANVAS");
                //canvas.width = w;
                //canvas.height = h;
                //// Draw image to canvas  
                //var context = canvas.getContext('2d');
                //context.drawImage(img[i], 0, 0, w, h);
                //// Get data URL encoding of image  
                //var uri = canvas.toDataURL("image/png");
                //$(img[i]).attr("src", img[i].src);
                //img[i].width = w;
                //img[i].height = h;
                //// Save encoded image to array  
                //images[i] = {
                //    type: uri.substring(uri.indexOf(":") + 1, uri.indexOf(";")),
                //    encoding: uri.substring(uri.indexOf(";") + 1, uri.indexOf(",")),
                //    location: $(img[i]).attr("src"),
                //    data: uri.substring(uri.indexOf(",") + 1)
                //};
            }

            // Prepare bottom of mhtml file with image data  
            var mhtmlBottom = "\n";
            for (var i = 0; i < images.length; i++) {
                mhtmlBottom += "--NEXT.ITEM-BOUNDARY\n";
                mhtmlBottom += "Content-Location: " + images[i].location + "\n";
                mhtmlBottom += "Content-Type: " + images[i].type + "\n";
                mhtmlBottom += "Content-Transfer-Encoding: " + images[i].encoding + "\n\n";
                mhtmlBottom += images[i].data + "\n\n";
            }
            mhtmlBottom += "--NEXT.ITEM-BOUNDARY--";

            //TODO: load css from included stylesheet  
            //var styles = "";
            var styles = '.header div p input{margin-left:5px;}.main {border:1px solid #000000;width:50%;height:100%;text-align:center;overflow-y:auto;margin-left:22%;}/*�����ı�����ʽ*/.mainContent {width:95%;height:400px;}/*������*/.table_title {margin:50px auto 30px auto;} /*�����ʽ*/.tab-content{height: 48px;line-height: 48px;border: 1px solid #33CCCC;font-size: 14px;border-collapse: collapse;width:95%;margin-left:2%;}/*tr  th  td�����ϱ߿�*/.tab-content tr{border:1px solid #33CCCC;height:63px;}.tab-content tr th{border:1px solid #33CCCC;}table tr td{border:1px solid #33CCCC;}/*��ť*/.btn{text-align: center;background-color: #33CCCC;width: 120px;height: 40px;color: white;}';
            // Aggregate parts of the file together  
            var fileContent = static.mhtml.top.replace("_html_", static.mhtml.head.replace("_styles_", styles) + static.mhtml.body.replace("_body_", markup.html())) + mhtmlBottom;

            // Create a Blob with the file contents  
            var blob = new Blob([fileContent], {
                type: "application/msword;charset=utf-8"
            });
            saveAs(blob, fileName + ".doc");
        };
    })(jQuery);
} else {123
    if (typeof jQuery === "undefined") {
        console.error("jQuery Word Export: missing dependency (jQuery)");
    }
    if (typeof saveAs === "undefined") {
        console.error("jQuery Word Export: missing dependency (FileSaver.js)");
    }
}