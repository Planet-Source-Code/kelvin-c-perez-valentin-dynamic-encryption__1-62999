'
' *************************************************************************
' *
' * Dymanic Encryption
' * Kelvin C. Pérez - Valentín
' * kelvin_perez@msn.com
' *
' * Encrypt text dynamically so the same text will never be encrypted equally
' * without the need to keep track of passwords (but can be modified to supply
' * the password separate from the main string).
' *
' * This is more like a character masking instead of a real ecnryption but will
' * do the trick.
' *
' * You can encrypt the same text as many times as you want and the resulting 
' * strings will always be different in size and characters.
' * 
' *************************************************************************
' *
' * Encryption Logic:
' * 
' *************************************************************************
' *
' *
1. create a random number from 1 to 99: encryption_Key

2. create a random number from 5 to 7: string_size

3. get the ASCII value of each character inside the text and multiply it 
   by the encryption_key 

4. Since the length of the resulting strings may very, equal all of them 
   to the same size, determined by string_size. Fill the necesary spaces 
   with random letters: A..Z and a..z

5. Mix the numbers with the letters without altering the original sequence 
   of the numbers.

6. Add all the sub-strings together to form the main string.

7. encrypt the "encryption_key" and "string_size" and insert them inside the 
   encrypted text.

8. Add some dummy strings (optional) to full people even more.
' * 
' *************************************************************************
' *
' * Decryption Logic:
' * 
' *************************************************************************
' *
1. Remove the Dummy strings

2. extract and decrypt the string_size in order to extract the encryption_key

3. extract and decrypt the encryption_key

4. extract each sub-string (the size of string_size).

5. Remove any letter from the string (get only the numbers).

6. Convert the resulting string to number (VAL() Function) and divide it by 
   the encryption_key

7. Convert the resulting number into it's character (CHR() Function) value.

8. Add all the characters together.
' * 
' *************************************************************************
' *
' * Credits:
' * 
' *************************************************************************
' *
1. OSEN excelent XP controls

2. The "About Form" & The Resize Sub I got it from PSC but can't find the 
   originals to give proper credits

3. Kevin Lawrence's Random Sequence (non-repetitive numbers). It can be found at:

gttp://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=892&lngWId=1

4. Last but not least to Heriberto Mantilla Santamaria for his support and feedback

You can find his PHP version of my code at:

http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=1760&lngWId=8

and an on-line demo at:

http://hackprotm.webcindario.com/EncryptDemo/demo.html

You can also encrypt the text with one version and decrypt it with the other (VB to PHP or PHP to VB)


The Code is not fully optimized since I started it as a tutorial for a friend.
