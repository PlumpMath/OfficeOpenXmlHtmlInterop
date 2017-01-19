
// var text = "<p><b>Welcome <i>Ke</i></b><i>nn</i><b><i>y</i></b></p>";
// var text = "<p>Hello World</p>";
//     var text ="<p style=\"text-align:right\"><u>Hello <i>World</i > < b > Kenny < / b > < / u > < / p > "
//     var text = "<p>Welcome<h1>Testing</h1></p>";
    var text = "<h1 style=\"text-align: center;\"><b><font color=\"OrangeRed\">Welcome</font> </b></h1><h2><b><i>Ke</i></b><i>nn</i><b><i>y</i></b></h2><p>Align to left</p><p style=\"text-align: right;\">align to right</p><p style=\"text-align: center;\">align to center</p>";
// var text = "<a href=\"http://www.google.com\" target=\"\">http://www.google.com</a>";


//var text = "<p><ul><li><b>a</b></li><li>b</li><li>c</li></ul></p>";

var output = testrun(text);

function testrun(text) {

    var output = "";
    var doc = new DOMParser().parseFromString(text, "text/html");
    console.log(doc)
    var bodyNode = doc.getElementsByTagName("body");

    console.log("Child Node Length", bodyNode[0].childNodes.length);

    for (var i = 0; i < bodyNode[0].childNodes.length; i++) {
        // console.log("Element ",i,":",bodyNode[0].childNodes[i]);
        var node = bodyNode[0].childNodes[i];
        console.log("Node", i, ":", node.nodeName);

        //does this node has child?
        if (node.childNodes.length > 0) {
            console.log("This node has Child Node");
            output += traverseInnerNode([], node);
        } else {
            console.log("This node does not have Child Node");
            output += printOutput('', node);
        }
        // output += "</w:p>";
    }

    // console.log(output);
    var n = document.getElementById("outputDiv");
    n.textContent = output;

}

//cater the case when it's order list / unordered list
function matchList(node) {

    var output = '';

    var pretag = '';
    pretag += "<w:pPr w:val=\"ListParagraph\">";

    //Handling the difference between ul and ol
    if (node.nodeName.toLowerCase() == 'ul') {
        pretag += " <w:numPr> <w:ilvl w:val=\"0\"/> <w:numId w:val=\"1\"/> </w:numPr> ";
    } else if (node.nodeName.toLowerCase() == 'ol') {
        pretag += " <w:numPr> <w:ilvl w:val=\"0\"/> <w:numId w:val=\"2\"/> </w:numPr> ";
    }

    pretag += "</w:pPr>";

    if (node.childNodes.length > 0) {
        var listNodes = node.childNodes;

        for (var i = 0; i < listNodes.length; i++) {


            var listOutput = '';
            listOutput += "<w:p>";
            listOutput += pretag;
            listOutput += traverseInnerNode([], listNodes[i]);
            listOutput += "</w:p>";

            output += listOutput;
        } //end of for-loop

    } //end of if

    return output;

}


function isParagraphStylingTag(node){
    if ('p' == node.tagName.toLowerCase() ||
        'h1' == node.tagName.toLowerCase() ||
        'h2' == node.tagName.toLowerCase()) {
        return true;
    }else{
        return false;
    }
}

function isTextRunStylngTag(node){
    if ('i' == node.tagName.toLowerCase() ||
        'u' == node.tagName.toLowerCase() ||
        'b' == node.tagName.toLowerCase() ||
        'font' == node.tagName.toLowerCase() ||
        'img' == node.tagName.toLowerCase()) {
        return true;
    }else{
        return false;
    }
}

//Traverse to the child node
/**
 * 2 types of node it will encounter,
 *   1. styling elements which goes to the w:rPr
 *      i, b, u, fonts, img, div...
 *   2. paragraph styling elements which goes to the w:pPr
 *      p, h1, h2, alignments
 *
 * @param outerTagHierarchy
 * @param node
 * @returns {string}
 */
function traverseInnerNode(outerTagHierarchy, node) {
    var tagHierarchy = outerTagHierarchy;

    var output = '';
    var outerTag = '';
    var singleTag = '';
    //If the currentTag is a p tag, then need to add a <w:p> to bound the follow

    //Start with a <w:p> if it matches those Paragraph styling tag
    if (isParagraphStylingTag(node)) {
        outerTag += "<w:p>";
    }
    outerTag += matchParagraphStyling(node);
    singleTag += matchSingleTagToNotation(node);
    tagHierarchy.push(singleTag);

    var innerNodes = node.childNodes;

    output += outerTag;
    for (var j = 0; j < innerNodes.length; j++) {

        var innerNode = innerNodes[j];

        if (innerNode.childNodes.length > 0) {
            // console.log("This inner node has Child Node");
            output += traverseInnerNode(tagHierarchy, innerNode);

        } else {

            var accumulatedTags = '';
            while (tagHierarchy.length) {
                accumulatedTags += tagHierarchy.pop();
                console.log("ACC TAGS ", accumulatedTags);
            }

            // console.log("This inner node does not have Child Node");
            output += printOutput(accumulatedTags, innerNode);
        }
    }

    //Close with a </w:p> if it matches those Paragraph styling tag
    if (isParagraphStylingTag(node)) {
        output += "</w:p>";
    }
    return output;
}

function printOutput(innerTag, node) {
    // console.log("In Print Output")
    var output = '';

    //Handle hyperlink tag
    if (node.parentNode.nodeName.toLowerCase() == 'a') {
        var hyperlink = '';
        hyperlink += "<w:hyperlink w:history=\"1\">";
        hyperlink += "<w:r>";
        hyperlink += "<w:rPr> ";
        hyperlink += "<w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" />";
        hyperlink += "<w:rStyle w:val=\"Hyperlink\" />";

        if (innerTag != '') {
            hyperlink += innerTag;
        }

        hyperlink += "</w:rPr>";
        hyperlink += "<w:t>" + node.textContent + "</w:t>";
        hyperlink += "</w:r>";
        hyperlink += "</w:hyperlink>";
        output += hyperlink;
    }
    //the normal case
    else {

        output += "<w:r>";
        if (node.nodeName == '#text') {
            if (innerTag != '') {
                output += "<w:rPr>";
                output += "<w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" />";
                output += innerTag;
                output += "</w:rPr>";
            }
            output += "<w:t xml:space=\"preserve\">" + node.textContent + "</w:t>";
        }

        // else if (node.nodeName.toLowerCase() == 'img'){
        //     var src = node.src;
        //     if (src != null && src != '') {
        //         output += "<pic:blipFill>";
        //         output += "<a:blip cstate=\"print\" link=\"" + src + "\"/>";
        //         output += "</pic:blipFill>";
        //     }
        // }

        //If the node is not a text node, style has to apply
        else {
            output += "<w:rPr>";
            if (innerTag != '') {
                output += innerTag;
            }
            output += matchSingleTagToNotation(node);
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
    //console.log(tagName);
    if (tagName != null) {
        switch (tagName.toLowerCase()) {
            case 'h1':
                paragraphStyle += "<w:pStyle w:val=\"Heading1\"/>";
                if (node.style != null && node.style.textAlign != null && node.style.textAlign != '') {
                    paragraphStyle += "<w:jc w:val=\"" + node.style.textAlign + "\"/>";
                }
                break;
            case 'h2':
                paragraphStyle += "<w:pStyle w:val=\"Heading2\"/>";
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

//Converting the HTML tag to the equivalent open office tag
//Single Tags
function matchSingleTagToNotation(node) {
    var output = '';
    var tagName = node.tagName;
    //console.log(tagName);
    if (tagName != null) {
        switch (tagName.toLowerCase()) {
            case 'i':
                output = "<w:i/>";
                break;
            case 'b':
                output = "<w:b/>";
                break;
            case 'u':
                output = "<w:u w:val=\"single\"/>";
                break;
            case 'font':
                //Currently only color is supported in the RichText Dialog box
                var color = node.color;
                if (color != null && color != '') {
                    output = "<w:color  w:val=\"" + color + "\"/>";
                }
                break;
            case 'div':
                console.log("Caught div");
                output = '';
                break;
            case 'img':
                console.log("Got img component");
                var src = node.src;
                if (src != null && src != '') {
                    output += "<pic:blipFill>";
                    output += "<a:blip cstate=\"print\" link=\"" + src + "\"/>";
                    output += "</pic:blipFill>";
                }
                break;
            default:
                output = "";
                break;
        } //end switch
    }
    return output;
}
