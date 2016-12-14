testrun();

function testrun() {
//     var text = "<p><b>Welcome <i>Ke</i></b><i>nn</i><b><i>y</i></b></p>";
//     var text = "<p>Hello World</p>";
//     var text ="<p style=\"text-align:right\"><u>Hello <i>World</i > < b > Kenny < / b > < / u > < / p > "
//     var text = "<p>Welcome<h1>Testing</h1></p>";
//     var text = "<h1 style=\"text-align: center;\"><b><font color=\"OrangeRed\">Welcome</font> </b></h1><h2><b><i>Ke</i></b><i>nn</i><b><i>y</i></b></h2><p>Align to left</p><p style=\"text-align: right;\">align to right</p><p style=\"text-align: center;\">align to center</p>\"";
// var text = "<a href=\"http://www.google.com\" target=\"\">http://www.google.com</a>";
var text = "<ul><li>a</li><li>b</li><li>c</li></ul>";
    var output = "";

    var doc = new DOMParser().parseFromString(text, "text/html");

    var bodyNode = doc.getElementsByTagName("body");


    console.log("Body Node", bodyNode);
    console.log("Body Children", bodyNode);
//----------------------------------------------------------------------------------------

    console.log("Child Node Length", bodyNode[0].childNodes.length);

    for (var i = 0; i < bodyNode[0].childNodes.length; i++) {
        // console.log("Element ",i,":",bodyNode[0].childNodes[i]);
        var node = bodyNode[0].childNodes[i];
        console.log("Node", i, ":", node.nodeName);

        if (node.nodeName.toLowerCase() == 'ul'){
            output += matchList(node);
        }
        else {
            output += "<w:p>";
            output += matchParagraphStyling(node);

            //does this node has child?
            if (node.childNodes.length > 0) {
                console.log("This node has Child Node")
                output += traverseInnerNode([], node);
            } else {
                console.log("This node does not have Child Node");
                output += printOutput('', node);
            }
            output += "</w:p>";
        }
    }
    console.log(output);
}

//cater the case when it's order list / unordered list
function matchList(node){
    var pretag = '';

    var output = "";
    output += "<w:p>";
    pretag += "<w:pPr w:val=\"ListParagraph\">";
    pretag += " <w:numPr> <w:ilvl w:val=\"0\"/> <w:numId w:val=\"1\"/> </w:numPr> ";
    pretag += "</w:pPr>";

    if (node.childNodes.length > 0){
        var listNodes = node.childNodes;
        for (var i=0; i<listNodes.length; i++){
            console.log("Traverse To Li Component");
            output += traverseInnerNode([pretag], listNodes[i]);
        }

    }

    output += "</w:p>";
    return output;

}

function traverseInnerNode(outerTagHierarchy, node) {
    var tagHierarchy = outerTagHierarchy;
    var outerTag = matchTagNameToNotation(node);
    tagHierarchy.push(outerTag);

    var innerNodes = node.childNodes;
    var output = '';

    for (var j = 0; j < innerNodes.length; j++) {

        var innerNode = innerNodes[j];

        if (innerNode.childNodes.length > 0) {
            // console.log("This inner node has Child Node");
            output += traverseInnerNode(tagHierarchy, innerNode);

        } else {

            //Todo: Get the loop of tag hierarchy out and append it to innerTag
            var accumulatedTags = '';
            while (tagHierarchy.length) {
                accumulatedTags += tagHierarchy.pop();
            }

            // console.log("This inner node does not have Child Node");
            output += printOutput(accumulatedTags, innerNode);
        }
    }
    return output;
}

function printOutput(innerTag, node) {
    // console.log("In Print Output")
    var output = '';

    //Handle hyperlink tag
    if (node.parentNode.nodeName.toLowerCase() == 'a'){
       var hyperlink = '';
       hyperlink += "<w:hyperlink w:history=\"1\">";
       hyperlink += "<w:r>";
       hyperlink += "<w:rPr> <w:rStyle w:val=\"Hyperlink\" /> </w:rPr>";
       hyperlink += "<w:t>"+node.textContent+"</w:t>";
       hyperlink += "</w:r>";
       hyperlink += "</w:hyperlink>";
       output += hyperlink;
    }

    // else if (node.parentNode.nodeName.toLowerCase() == 'li'){
    //     output += innerTag
    // }

    //the normal case
    else{

        output += "<w:r>";
        if (node.nodeName == '#text') {
            if (innerTag != '') {
                output += "<w:rPr>";
                output += innerTag;
                output += "</w:rPr>";
            }
            output += "<w:t xml:space=\"preserve\">" + node.textContent + "</w:t>";
        }
        //If the node is not a text node, style has to apply
        else {
            output += "<w:rPr>";
            if (innerTag != '') {
                output += innerTag;
            }
            output += matchTagNameToNotation(node);
            output += "</w:rPr>";
            output += "<w:t xml:space=\"preserve\">" + node.textContent + "</w:t>";
        }
        output += "</w:r>";
    }


    return output;
}

function matchParagraphStyling(node) {

    var output = '';
    var paragraphStyle = '';
    var tagName = node.tagName;
    console.log(tagName);
    if (tagName != null) {
        switch (tagName.toLowerCase()) {
            case 'h1':
                paragraphStyle += "<w:pStyle w:val=\"Heading1\"/>"
                if (node.style != null && node.style.textAlign != null && node.style.textAlign != '') {
                    paragraphStyle += "<w:jc w:val=\"" + node.style.textAlign + "\"/>";
                }
                break;
            case 'h2':
                paragraphStyle += "<w:pStyle w:val=\"Heading2\"/>"
                if (node.style != null && node.style.textAlign != null && node.style.textAlign != '') {
                    paragraphStyle += "<w:jc w:val=\"" + node.style.textAlign + "\"/>";
                }
                break;
            case 'p':
                if (node.style != null && node.style.textAlign != null && node.style.textAlign != '') {
                    paragraphStyle = "<w:jc w:val=\"" + node.style.textAlign + "\"/>";
                }
                break;
        }//end of switch
        if (paragraphStyle != '') {
            output += "<w:pPr>";
            output += paragraphStyle;
            output += "</w:pPr>"
        }
    }

    return output;
}

function matchTagNameToNotation(node) {
    var output = '';
    var tagName = node.tagName;
    console.log(tagName);
    if (tagName != null) {
        switch (tagName.toLowerCase()) {
            case 'i':
                output = "<w:i/>";
                break;
            case 'b':
                output = "<w:b/>";
                break;
            case 'u':
                output = "<w:u/>";
                break;
            case 'font':
                //Currently only color is supported in the RichText Dialog box
                var color = node.color;
                output = "<w:color  w:val=\"" + color + "\"/>"
                break;
            case 'img':
                console.log("Got img component");
                output = "";
                break;
            default:
                output = "";
                break;
        } //end switch
    }
    return output;
}
