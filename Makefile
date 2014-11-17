ZLIB=\
zlib/adler32.c\
zlib/compress.c\
zlib/crc32.c\
zlib/deflate.c\
zlib/gzclose.c\
zlib/gzlib.c\
zlib/gzread.c\
zlib/gzwrite.c\
zlib/infback.c\
zlib/inffast.c\
zlib/inflate.c\
zlib/inftrees.c\
zlib/trees.c\
zlib/uncompr.c\
zlib/zutil.c

MINIZIP=\
minizip/ioapi.c\
minizip/iowin32.c\
minizip/mztools.c\
minizip/unzip.c\
minizip/zip.c

CFLAGS = -O2 -Wall

all : zip.dll base64.dll winapi.dll

zip.dll : luazip.c $(ZLIB) $(MINIZIP)
	gcc $(CFLAGS) --shared -o $@ $^ -Izlib -Iminizip -I/usr/local/include -L/usr/local/bin -llua52

base64.dll : lbase64.c
	gcc $(CFLAGS) --shared -o $@ $^ -I/usr/local/include -L/usr/local/bin -llua52

winapi.dll : winapi.c
	gcc $(CFLAGS) --shared -o $@ $^ -I/usr/local/include -L/usr/local/bin -llua52 -luser32 -lole32

clean :
	rm zip.dll base64.dll winapi.dll