JFW Script File                                                          �  h    $scriptfilename    1  The MuseScore scripts are running
The application name is %1
This is version %2 of the scripts
Release date: %3      getappfilename  1 ScriptVersion1.0.1  1 ReleaseDateFebruary 11, 2020      formatstring    '        issamescript            (   %     saymessage                %     saymessage           �     windowactivatedevent          %     getwindowname   '  %    MuseScore3  
        %     gettypeandtextstringsforwindow  '          %    saymessage     	         %     windowactivatedevent              focuschangedeventex               getobjectrole   '  %    
   
             getobjectname          say    	      %   %  
  # � %  %  
  
  # � %  %  
  
     	         %   %  %  %  %  %  %    focuschangedeventex       �     valuechangedevent            %        %    braillestring         %         say    	         %   %  %  %  %  %  %    valuechangedevent         l     namechangedevent               %   %  %  %  %  %    namechangedevent          8     $passkey         typecurrentscriptkey          