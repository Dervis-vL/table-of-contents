def paragraph_sort(dictionary):
    for chap, pars in dictionary.items():
        if len(pars) == 0:
            continue
        elif len(pars[0].split(".")) == 2:
            pars.sort(key=lambda x: int(x.split(".", 1)[1].split(" ", 1)[0]))
        elif len(pars[0].split(".")) == 3:
            pars.sort(key=lambda x: int(x.split(".")[1]))
    
    return dictionary



if __name__ == "__main__":
    sample_dict = {}

    paragraph_sort(sample_dict)