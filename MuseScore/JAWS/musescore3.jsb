JFW Script File                                                          �  P    $scriptfilename     The MuseScore scripts are running
The application name is %1
This is version %2 of the scripts
Release date: %3      getappfilename   1.0.1    February 11, 2020     formatstring    '        issamescript            (   %     saymessage                %     saymessage           �     windowactivatedevent          %     getwindowname   '  %   MuseScore3  
        %     gettypeandtextstringsforwindow  '          %    saymessage     	         %     windowactivatedevent          H    focuschangedeventex                     getcurrentobject    '  %            accrole '  %    
   
             getobjectname          say    	      %   %  
  # � %  %  
  
  # � %  %  
  
     	         %   %  %  %  %  %  %    focuschangedeventex       �     valuechangedevent            %        %    braillestring         %         say    	         %   %  %  %  %  %  %    valuechangedevent         l     namechangedevent               %   %  %  %  %  %    namechangedevent          8     $passkey         typecurrentscriptkey          