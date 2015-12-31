#include <stdio.h>
#include <stdint.h>
#include <string.h>

uint32_t jhash(char *key, size_t len)
{
    uint32_t hash, i;
    for(hash = i = 0; i < len; ++i)
    {
        hash += key[i];
        hash += (hash << 10);
        hash ^= (hash >> 6);
    }
    hash += (hash << 3);
    hash ^= (hash >> 11);
    hash += (hash << 15);
    return hash;
}

int main(int argc, char *argv[]) {
    if (argc != 2) {
      printf("Usage: jenkins <my string>\n");
    }
    printf("%s: %i\n", argv[1], jhash(argv[1], strlen(argv[1])));
    return 0;
}
