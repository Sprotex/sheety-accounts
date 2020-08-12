
def read_config():
    with open("config.txt") as file:
        lines = file.read().splitlines()

    conf = {}
    for line in lines:
        split = line.split("=")
        assert(len(split) == 2)
        conf[split[0]] = split[1]
    return conf


if __name__ == '__main__':
    print(read_config())
