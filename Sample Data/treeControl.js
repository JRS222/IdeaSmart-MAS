
//treeview CSS classes
var CLASSES = ($.treeview.classes = {
    open: "open",
    closed: "closed",
    expandable: "expandable",
    expandableHitarea: "expandable-hitarea",
    lastExpandableHitarea: "lastExpandable-hitarea",
    collapsable: "collapsable",
    collapsableHitarea: "collapsable-hitarea",
    lastCollapsableHitarea: "lastCollapsable-hitarea",
    lastCollapsable: "lastCollapsable",
    lastExpandable: "lastExpandable",
    last: "last",
    hitarea: "hitarea"
});
        

function ini_tree()
{
    $("#phbk_tree").treeview({
        animated: "fast",
        collapsed: true
    });
    initial_li_div_task();
}
function initial_li_div_task()
{
    $("#phbk_tree li span").each(function(){
        $(this).click(function(event){                                
            event.stopPropagation();  
            var _li=$(this).parent(); 
            if(_li.attr("sno") == "1")//section 1
            {
                $("#section_1_content").show()     
                $("#ie7fix").hide();
                _li.find("ul").removeClass("placeholder");
                if($("#section_1_content #My_Check_Point").length == 0)
                {          
                    $("#section_1_content").load("content/section_1.php?msbookno="+booknum+"&volno="+vol, function(){
                    });        
                } 
            }
            else
            {
                if(_li.attr("loaded")!="y")
                {
                    _li.find(">ul").addClass("placeholder");
                    _li.attr("loaded","y");                        
                    ready_to_load_branch(_li);
                    //also need to load main page if this is topic
                }
                do_toggle(_li);                    
            }
        });
        $(this).hover(function(){
            $(this).css("color","red");                    
        },
        function(){
            $(this).css("color","navy");
        });
    });            
    
    $("#phbk_tree li div").click(function(event){
        event.stopPropagation();            
        var _li=$(this).parent();
        if(_li.attr("sno") == "1")//section 1
        {
            $("#section_1_content").show()     
            $("#ie7fix").hide();
            _li.find("ul").removeClass("placeholder");
            if($("#section_1_content #My_Check_Point").length == 0)
            {          
                $("#section_1_content").load("content/section_1.php?msbookno="+booknum+"&volno="+vol, function(){
                });        
            } 
        }
        else
        {
            if(_li.attr("loaded")!="y")
            {
                _li.find(">ul").addClass("placeholder");
                _li.attr("loaded","y");
                ready_to_load_branch(_li);
            }
            do_toggle(_li);
        }
    });
}
//toggle event
function do_toggle(_li)
{
    _li.find(">.hitarea")
        .swapClass( CLASSES.collapsableHitarea, CLASSES.expandableHitarea )
        .swapClass( CLASSES.lastCollapsableHitarea, CLASSES.lastExpandableHitarea )
    .end()
        // swap classes for parent li
        .swapClass( CLASSES.collapsable, CLASSES.expandable )
        .swapClass( CLASSES.lastCollapsable, CLASSES.lastExpandable )
        .find( ">ul" ).toggle("fast");
}

//load leaf
function ready_to_load_branch(_li)
{   
    var sec_num = _li.attr("sno");
    var sub_ul = _li.find("ul");   
    var special_char = _li.attr("special_char");
    if(sec_num == 1)
    {
        //now we are ready to load static html page here     
        $("#section_1_content").show()     
        $("#ie7fix").hide();
        sub_ul.removeClass("placeholder");
        if($("#section_1_content #My_Check_Point").length == 0)
        {          
            $("#section_1_content").load("content/section_1.php?msbookno="+booknum+"&volno="+vol, function(){
            });        
        } 
        return;
    }
    else
    {
        $("#ie7fix").show();
        $("#section_1_content").hide();
    }
    $.ajax({
        url: "ajax/section_figures.php",
        type:"POST",
        data:{booknum:booknum, vol:vol, section:sec_num},
        success: function(data){                        
            $(data).appendTo(sub_ul);
            //$("#phbk_tree").treeview({add:data}); 
            sub_ul.removeClass("placeholder");
            var firstfigno = sec_num + "-1";
            var figitemno = 0;
            if(special_char != sec_num)
                firstfigno = special_char + "-1";
            if (sno != "" && figno != "" && itemno != "" && para_fig_load_ctn == 0){  
                firstfigno = figno;  
                figitemno = itemno;
                para_fig_load_ctn = para_fig_load_ctn + 1;              
                $("div.left_btm").animate({
                    scrollTop: $("li[figno='"+figno+"']").offset().top
                }, 2000);
            }      
            load_figure(firstfigno);
            load_figure_table(firstfigno,figitemno);
      }
    });
}
