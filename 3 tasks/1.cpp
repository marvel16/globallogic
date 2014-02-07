// Разверните строку. Указатель reverse_string должен 
// указывать на развернутую строку.

char* string = “The string!”;

int main()
{
	char* reverse_string;
	int Len = strlen(string); 
	int midLen = Len/2;// creating temp so it doesnt calculate strlen/2 in cycle all the time
	
	for(int i=0; i<midLen; i++)
	{	
		char chBuf = *(string+i);
		*(string+i) = *(string+Len-i);
		*(string+len-i) = chBuf;
	}
	reverse_string=string;
	return 0;
}