import matplotlib.pyplot as plot
import requests

data=requests.get("https://pomber.github.io/covid19/timeseries.json")
jsonoutput=data.json()
fig,plots=plot.subplots(2)
plot.subplots_adjust(hspace=0.5)


def getcountrycnfcase(name):
    xaxis=list()
    yaxis=list()
    arraylen=len(jsonoutput[name])
    maxval=arraylen
    minval=arraylen-7
    for dates in range(minval,maxval):
        xaxis.append(jsonoutput[name][dates].get("date",""))
        yaxis.append(jsonoutput[name][dates].get("confirmed",""))
    xaxis=list(dict.fromkeys(xaxis))
    yaxis=list(dict.fromkeys(yaxis)) 
    plots[0].plot(xaxis,yaxis,label=name)
    
    plots[0].set_title("Confirmed Covid 19 Case in last 7 days")
    
    
def getcountrycnfdeaths(name):
    xaxis=list()
    yaxis=list()
    arraylen=len(jsonoutput[name])
    maxval=arraylen
    minval=arraylen-7
    for dates in range(minval,maxval):
        xaxis.append(jsonoutput[name][dates].get("date",""))
        yaxis.append(jsonoutput[name][dates].get("deaths",""))
    xaxis=list(dict.fromkeys(xaxis))
    yaxis=list(dict.fromkeys(yaxis)) 
    plots[1].plot(xaxis,yaxis,label=name)

    plots[1].set_title("Covid 19 deaths in last 7 days")
   


getcountrycnfcase("Italy")
getcountrycnfcase("US")     
getcountrycnfcase("Spain")     
getcountrycnfcase("France")  
getcountrycnfcase("Iran")
getcountrycnfcase("China")



getcountrycnfdeaths("Italy")
getcountrycnfdeaths("US")     
getcountrycnfdeaths("Spain")     
getcountrycnfdeaths("France")  
getcountrycnfdeaths("Iran")
getcountrycnfdeaths("China")


plot.legend()
plot.show()