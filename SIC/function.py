import pandas as pd
import numpy as np

data = pd.read_excel('TMI_IV_versi_Excel.xlsx')

def max_age():
    max_age = data.iloc[-1,0]
    return max_age

def int_probability(age, interval, sex):
    result = 1
    
    if sex == 'male':
        for i in range(interval):
            y = float(data[data.iloc[:, 0] == age + i][data.columns[1]].iloc[0])
            result *= (1 - y)
    else:
        for i in range(interval):
            y = float(data[data.iloc[:, 0] == age + i][data.columns[2]].iloc[0])
            result *= (1 - y)
            
    result2 = 1 - result

    return round(result,8),round(result2,8)

def UDD(age,interval,sex):
    
    if type(age) == 'int':
        tqx = int_probability(age,1,sex)[1] * interval
        tpx = 1 - tqx

    else : 
        x = divmod(age,1)
        tqx = (interval * int_probability(x[0],1,sex)[1]) / (1-(x[1]*int_probability(x[0],1,sex)[1]))
        tpx = 1 - tqx

    return tpx,tqx

def CF(age,interval,sex):
    x = divmod(age,1)

    tpx = (int_probability(x[0],1,sex)[0]) ** interval
    tqx = 1 - tpx

    return tpx,tqx

def Hyperbolic(age,interval,sex):
    
    if type(age) == 'int':
        tqx = (interval * int_probability(age,1,sex)[1]) / (1-((1-interval)*int_probability(age,1,sex)[1]))
        tpx = 1 - tqx

    else :
        x = divmod(age,1)
        tqx = (interval * int_probability(x[0],1,sex)[1]) / (1-((1-interval-x[1])*int_probability(x[0],1,sex)[1]))
        tpx = 1 - tqx

    return tpx,tqx

def get_probability(age,interval,sex,assumption):

    not_fractional = (divmod(age,1)[1] == 0) and (divmod(interval,1)[1]==0)

    if not_fractional == True :
        return int_probability(age,interval,sex)
    
    else :
        switch = {
            'UDD' : UDD,
            'CF' : CF,
            'Hyperbolic' : Hyperbolic
        }
        tpx = 1

        age_array = np.arange(age,age+interval,1)
        int_array = np.array([])

        if interval < 1 :
            int_array = np.array([interval])
        else :
            int_array = np.ones(int(interval))

            if (divmod(age,1)[1] + divmod(interval,1)[1]) <= 1 :
                int_array = np.append(int_array,divmod(interval,1)[1])

            else :
                remainder1 = 1 - divmod(age,1)[1]
                remainder2 = divmod(interval,1)[1] - remainder1
                int_array = np.append(int_array,[remainder1,remainder2])
                age_array = np.append(age_array,age_array[-1]+remainder1)

        for i,j in zip (age_array,int_array):
            tpx = tpx * switch[assumption](i,j,sex)[0]
        
        tqx = 1- tpx

        return tpx,tqx  
    
def annuity(age,sex,duration,periodic,assumption,interest_rate):

    n = divmod(duration/periodic,1)[1]

    if n == 0:
        n_array = np.arange(0,duration,periodic)
        probability_array = np.array([])
        discount_factor_array = np.array([])

        for i,j in zip(n_array,interest_rate):
            probability_array = np.append(probability_array,get_probability(age,i,sex,assumption)[0])
            discount_factor_array = np.append(discount_factor_array,(1+j)**-i)

        result = np.dot(probability_array,discount_factor_array)

        return result
    
    else :
        raise ValueError('Annuities are not integer type')
    

def wholelife(age,interval,sex,deferred,assumption,interest_rate,benefit):

    max_age = data.iloc[-1,0]
    age_interval = max_age - age
    n_array = np.arange(0,age_interval,interval)
    np.append(n_array,age_interval)

    n_array2 = n_array + age
    life_probability = np.array([])
    death_probability = np.array([])
    discount_factor_array = np.array([])

    result = 0

    for i,j,k in zip(n_array,n_array2,interest_rate):
        life_probability = np.append(life_probability,get_probability(age,i,sex,assumption)[0])
        death_probability = np.append(death_probability,get_probability(j,interval,sex,assumption)[1])
        discount_factor_array = np.append(discount_factor_array,(1+k)**(-i-interval))

    for i,j,k in zip(life_probability,death_probability,discount_factor_array):
        result = result + (i*j*k)


    result2 = 0
    defferred_interval = int(deferred/interval)

    if deferred == 0 :
        return benefit * result
    else :
        for i in range(0,defferred_interval) :
            result2 = result2 + (discount_factor_array[i]*life_probability[i]*death_probability[i])

        return benefit * (result - result2)
    
def ntermslife(age,interval,sex,nterms,deferred,assumption,interest_rate,benefit) :

    life_probability = np.array([])
    death_probability = np.array([])
    discount_factor_array = np.array([])
    result = 0
    
    if deferred == 0 :
        n_array = np.arange(0,nterms,interval)
        n_array2 = n_array + age

        for i,j,k in zip(n_array,n_array2,interest_rate):
            life_probability = np.append(life_probability,get_probability(age,i,sex,assumption)[0])
            death_probability = np.append(death_probability,get_probability(j,interval,sex,assumption)[1])
            discount_factor_array = np.append(discount_factor_array,(1+k)**(-i-interval))
        
        for i,j,k in zip(life_probability,death_probability,discount_factor_array):
            result = result + (i*j*k)

        return result * benefit
    
    else :
        n_array = np.arange(0,nterms,interval)
        n_array = n_array[int(deferred/interval):]
        np.append(n_array,n_array[-1]+interval)

        n_array2 = n_array + age

        for i,j,k in zip(n_array,n_array2,interest_rate):
            life_probability = np.append(life_probability,get_probability(age,i,sex,assumption)[0])
            death_probability = np.append(death_probability,get_probability(j,interval,sex,assumption)[1])
            discount_factor_array = np.append(discount_factor_array,(1+k)**(-i-interval))
        
        for i,j,k in zip(life_probability,death_probability,discount_factor_array):
            result = result + (i*j*k)

        return result * benefit
    
def pure_endowment(age,interval,sex,interest_rate,assumption,benefit):

    result = ((1+interest_rate)**(-interval)) * get_probability(age,interval,sex,assumption)[0]
    return result * benefit

def endowment(age,interval,sex,interest_rate1,interest_rate2,n_terms,assumption,benefit1,benefit2):

    result1 = ntermslife(age,interval,sex,n_terms,0,assumption,interest_rate1,benefit1)
    result2 = pure_endowment(age,n_terms,sex,interest_rate2,assumption,benefit2)

    return result1+result2

class insurance :

    def __init__(self,age,interval,sex,assumption):
        self.age = age
        self.interval = interval
        self.sex = sex
        self.assumption = assumption

    def wholelife(self,deferred,interest_rate,benefit):
        return wholelife(self.age,self.interval,self.sex,deferred,self.assumption,interest_rate,benefit)
    
    def ntermslife(self,nterms,deferred,interest_rate,benefit):
        return ntermslife(self.age,self.interval,self.sex,nterms,deferred,self.assumption,interest_rate,benefit)
    
    def pure_endowment(self,nterms,interest_rate,benefit):
        return pure_endowment(self.age, nterms, self.sex, interest_rate, self.assumption, benefit)
    
    def endowment(self,interest_rate1,interest_rate2,n_terms,benefit1,benefit2):
        return endowment(self.age,self.interval,self.sex,interest_rate1,interest_rate2,n_terms,self.assumption,benefit1,benefit2)
    
def is_number(value):
    try:
        float_number = float(value)
        return True
    except ValueError:
        return False