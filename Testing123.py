

def test(i):
    return testing(i)

def testing(j):
    return j+1

def main():
    print(test(1))
    strings = "2019-02-20"
    print(strings[:4])
    diction = {}
    diction["Year 1"] = {}
    diction["Year 1"]["Revenue"] = "8400"
    print(diction["Year 1"]["Revenue"])

if __name__ == "__main__":
    main()
