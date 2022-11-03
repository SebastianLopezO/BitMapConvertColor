Algoritmo sin_titulo
	Definir Max, Min,dif Como Entero

	Dimension vector[3]
	leer vector[0]
	leer vector[1]
	leer vector[2]
	Max=-1
	Min=256
	Para i=0 Hasta 2 Con Paso 1 Hacer
		Si vector[i] Es Mayor Que Max Entonces
			Max=i
		FinSi
		si vector[i] Es Menor Que Min Entonces
			Min=i
		FinSi
	FinPara
		dif= vector[Max]-vector[Min]
		NTSC=vector[Min]+(dif/2)	
	Escribir NTSC
	
FinAlgoritmo
