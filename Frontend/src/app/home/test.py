t = "JNTFuchsia Fedora651"

First = t[-3:]
sec = t[:3]

t = t[3:-3]

print(f"init {First}")
print(f"sec {sec}")
print("Final", t)
