(function($){  
    $.fn.extend({  
        setTime: function(){  
            var $this = $(this);  
              
            $(this).click(function(){  
                $("div").detach("#settime");  
                $("body").append(createDiv());  
                $("#settime").css({  
                    padding:"3px",  
                    position :"absolute",  
                    "background-color":"#E3E3E3",   
                    "font-size": "12px",   
                    border: "1px solid #777777",  
                    width:"170px",  
                    top: $this.offset().top + 22,  
                    left: $this.offset().left  
                })  
                  
                $("input#time_submit").click(function(){  
                    $this.val($("select[name='hour']").val() + ":" + $("select[name='minute']").val());  
                    $("div").detach("#settime");  
                })  
            })  
              
            function createDiv(){  
                var str="<div id='settime' style='z-index:999'> 时<select name='hour'>";   
                for (h = 0; h <= 9; h++) {   
                    str += "<option value='0" + h + "'>0" + h + "</option>";   
                }   
                for (h = 10; h <= 23; h++) {   
                    str += "<option value='" + h + "'>" + h + "</option>";   
                }   
                str += "</select> 分<select name='minute'>";   
                for (h = 0; h <= 9; h++) {   
                    str += "<option value='0" + h + "'>0" + h + "</option>";   
                }   
                for (h = 10; h <= 59; h++) {   
                    str += "<option value='" + h + "'>" + h + "</option>";   
                }  
                str += "</select> <input id='time_submit' type='button' value='确定' style='font-size:12px' mce_style='font-size:12px'/></div>";  
                return str;  
            }  
              
        }  
    })  
})(jQuery);