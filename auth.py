import hmac, hashlib, binascii

# 由你提供的 PBKDF2-SHA256 帳密資料庫
PWD_DB = {
    "Charles": {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "32ae892164a22af5f83261bd239ed304", "hash": "27fb5fb7bbe2629d8c53dbbdf021423cdb4e7015e5858deafb3a0e405139bb40"},
    "Hsiang":  {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "9c31cdf98b82aa1741154680e456e3e0", "hash": "292e30442d243ea5f82879f1ce71f9ff2dc600f7234a075ba3ee130f45eb29b4"},
    "Sandy":   {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "dd8fbb2a735b076e5cff3bdee67fc3cf", "hash": "7bff0a1388c1447e934552175786d2fa5b9bc9b17ac3d9da246182dd7ec31e35"},
    "Min":     {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "a4a89474d39a1d89ac652a56ccd33301", "hash": "7d788d76be27923209c08aba44fdfc0ca6ce5530ed4b91283810fd0c34bc1a0f"},
    "May":     {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "88d33f6eb3d9a6506b705c3810e7be0b", "hash": "53765f6d56af8c2e49f917c89d60212ab8aeec28d215c9e53cf394e897782631"},
    "Ping":    {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "4af5ee4403ad13cb6a2b0836da5d02b1", "hash": "1c1757b927959d2ef8897467f1c823753ec166f0d5c0a1a8ed5d91a84f2ab00d"},
    "Denny":   {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "bc88ba930b619a25dcce81e6ee616305", "hash": "3dfe81a7dd31acaf2816604c000637f328049d1ca9f13940e217ec51f3a5e7c7"},
    "Davina":  {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "8ce1cb7106316a21db1b48534d7d1833", "hash": "3a79b1feaa96cd7dc7dbced0bc2226d84da22ecda5a38d7d44a58f98e8c24b96"},
    "Arthur":  {"algo": "pbkdf2_sha256", "iter": 200000, "salt": "8e9a0b3e6c6dd1dccd6964101b5af752", "hash": "0409292dedb20de507c7fae67d25f502998c80cb4fcace6758d8fedc042d5570"},
}

def pbkdf2_sha256(password: str, salt_hex: str, iter: int, dklen: int = 32) -> str:
    salt = binascii.unhexlify(salt_hex)
    dk = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, iter, dklen=dklen)
    return binascii.hexlify(dk).decode('ascii')

def verify(username: str, password: str) -> bool:
    rec = PWD_DB.get(username)
    if not rec or rec.get('algo') != 'pbkdf2_sha256':
        return False
    expect_hex = rec['hash']
    got_hex = pbkdf2_sha256(password, rec['salt'], rec['iter'])
    # 常數時間比較，避免時序攻擊
    return hmac.compare_digest(expect_hex, got_hex)
