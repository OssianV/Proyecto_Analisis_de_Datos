# Proyecto Analisis de Datos

---

Abstract: [PENDIENTE] - Poner un resumen del proyecto de un parrafo y pocas oraciones.

> NOTA - Por el momento, la organizacion del proyecto la llevaremos a cabo en este READ.me. Al finalizar el proyecto ya lo convertimos en un READ.me normal.

## Planificacion del proyecto y trackeo de responsabilidades individuales

En esta seccion vamos a planear que es lo que se tiene que hacer, ademas de asignar tareas individuales a cada quien. Usaremos checkboxes de markdown para llevar control de que tareas se han terminado y cuales siguien pendientes. Para marcar una como completado, simplemente agreguen una "x" dentro de los corchetes. Ademas, se agregra una seccion de implementaciones **extra**, que estaria chido agregar al proyecto, aunque no son escenciales.

### Tareas

Lo primero que tenemos que hacer es entender la base de datos; saber exactamente que informacion contiene. Para esto, hay que analisar cada una de las tablas que estan contenidas en ella. Esto implica saber que es lo que se mide en cada una de las columnas; cantidad de tiros? cantidad de goles? cantidad de pases? alguna estadistica mas rebuscada?

Es importante definir cada una de estas. Como este trabajo es colaborativo, habra que registrar el analisis para que los demas podamos entender. Esto se hara en la [Libreta_resumen_bd](/SCRIPTS/Libreta_resumen_bd.ipynb). Como la base de datos tiene 7 tablas, se me ocurrio dividir esto en dos:

- Documentar tabla `Match` - 115 columnas
- Documentar las tablas:
    - `Player` - 7
    - `Player_Attributes` - 42
    - `Team` - 5
    - `Team_Attributes` - 25
    - `Country` - 2
    - `League`- 3

Esto se hara entre dos personas; idealmente los que mas familiaridad tengan con el deporte.

Por el momento, parece que la base de datos es lo suficientemente extensa para escoger distintas direcciones para nuestro analisis. Esto es, podemos enfocarnos en algunos aspectos concretos y desarrollar todo el analisis sobre ello. Por ejemplo, enfocarnos en hacer un analisis de jugadores, asi omitiendo la perspectiva de los equipos. 

Debido a la variedad de posibilidades, cada quien debera plantear una direccion de como dirigir el analisis. Enfocarse en estadisticas de resultados por equipo? Enfocarse en una liga en concreto? Enfocarse en jugadores? Centrarnos en algunos atributos de los jugadores? Buscar relaciones entre la efectividad de los jugadores y los equipos en los que juegan? - Preguntas dirigidas, pero cuya respuesta no la de una grafica o estadistico en concreto.

Cada quien planteara uno de estos analisis. Ya veremos si hacemos uno, varios u todos.

Una vez que definamos el/los enfoque(s) de nuestro(s) analisis, podremos proceder a completar la [Libreta_resumen_bd](/SCRIPTS/Libreta_resumen_bd.ipynb) al especificar que columnas son de nuestro interes o no. Esto con el objetivo de modificar las queries que hagamos a la base de datos, asi solo traernos la informacion que nos sea relevante.

A partir de aqui, lo que hagamos dependera de el enfoque del analisis, por lo cual ya no podemos planear por adelantado.

#### Direccion del analisis

> Ossian

Se me ocurrio plantear una direccion de analisis donde nos enfocaremos en obtener insights relevantes que podamos dar a un director tecnico de una partido. 

Nos centraremos en analizar y generar sugerencias del tipo de jugador a scouter, de las formacion a usar, de las mejores formas de identificar un jugador valioso, de los atributos de jugadores que mayor impacto tienen en ganar un partido, etc. El analisis entonces podra incluir observaciones hechas a travez de las diferentes ligas que se incluyen en el dataset.

#### Ideas para analisis

Hay que hacer minimo 8 graficos diferentes, a continuacion ponemos una lista de algunas opciones (marquen adecuadamente si ya se hizo una isntancia de dicho tipo de grafico)

- [ ] Heatmap
- [ ] Bubble map
- [ ] Bar chart
- [ ] Stacked bar chart
- [ ] Pie chart
- [ ] Ring chart
- [ ] Boxplot
- [ ] Time series
- [ ] Histogram
- [ ] Violin plot
- [ ] Waffle plot
- [ ] Bubble plot
- [ ] Styled table
- [ ] Scatter plot
- [ ] Circular barplot
- [ ] Radar chart

Plantee 20 ideas de analisis que pueden hacer, y los graficos que se usarian para responderlo. Pueden checarlas [aqui](Libreta_planteamiento_de_insights.ipynb). Leeanlo y seleccionen acordemente (recuerden que debe hacer 8 tipos de graficos diferentes minimo).

### Extras

Estaria interesante hacer mas de lo necesario. Es por eso que aqui pondre ideas que estaria chido llevar a cabo.

> Notese que nada de esto fue solicitado por el profe

- [ ] Crear un dashboard ejecutivo.
- [ ] Hacer analisis usando algoritmos de machine learning basicos (regresion lineal, clasficacion, arboles de desicion, etc.)
- [ ] Crear una presentacion ejecutiva de nuestros insigths.

### Asignacion de tareas

Ossian:

- [x] Plantear una direccion del analisis.
- [x] Documentar tabla `Match`.
- [x] Documentar las tablas:
        - `Player`
        - `Player_Attributes`
        - `Team`
        - `Team_Attributes`
        - `Country`
        - `League`
- [ ] Preparar los tres dataframes que se usaran para la generacion de graficos
    - [ ] df_player_latest
    - [ ] df_match_base
    - [ ] df_tean_match
- [ ] Hacer los 4 analisis que me tocan:
    - [ ] 
    - [ ]
    - [ ]
    - [ ]

Estefania:

- [ ] Hacer graficos para 2 insights (de preferencia 3 pls)

Ian:

- [ ] Hacer graficos para 2 insights (de preferencia 3 pls)

Karim:

- [ ] Hacer graficos para 2 insights (de preferencia 3 pls)

## Vinculos a archivos importantes dentro de repositporio

Aqui pondre algunos hipervinculos a archivos importantes dentro del repositorio

- [Libreta_sqlite_basics](SCRIPTS/Libreta_sqlite_basics.ipynb): Libreta con un tutorial basico de como leer las tablas del archivo .sqlite y pasarlas a un dataframe + un poco de teoria
- [Libreta_resumen_bd](SCRIPTS/Libreta_resumen_bd.ipynb): Libreta resumen de la base de datos.

## Tutoriales, articulos, documentacion relevante

- La base de datos en kaggle es [esta](https://www.kaggle.com/datasets/hugomathien/soccer)
- Si no saben usar github y git, yo aprendi lo superbasico usando [esta lista de reproduccion](https://www.youtube.com/watch?v=BCQHnlnPusY&list=PLRqwX-V7Uu6ZF9C0YMKuns9sLDzK6zoiV). Alternativamente, pueden usar los propios [tutoriales de github](https://docs.github.com/en/get-started/start-your-journey/hello-world). O solo preguntenle a un LLM.
- Para que aprendan a usar la funcion de "issues", solo vean [este video](https://www.youtube.com/watch?v=WMykv2ZMyEQ&list=PLRqwX-V7Uu6ZF9C0YMKuns9sLDzK6zoiV&index=4)
- Como la base de datos es un archivo .sqlite, para aprender lo basico de sqlite yo estoy usando [este recurso](https://www.sqlitetutorial.net/)


Extras:
- Lo basico de pandas lo pueden encontrar [aqui](https://pandas.pydata.org/docs/user_guide/index.html)
- Para que entiendan como funciona la programacion modular vean [este short](https://www.youtube.com/shorts/Ju6tP03GI7c). Para que vean como se veria un modulo creado por ustedes muy sencillo vean [este video](https://www.youtube.com/watch?v=cgxEqlGJcrY). Cuando yo he usado modulos siempre defino clases, y funciones dentro de clases; si puueden hagan lo mismo, aunque como se ve en el ultimo video esto no es necesario.
