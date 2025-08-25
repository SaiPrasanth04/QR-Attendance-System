import pandas as pd
import os
from werkzeug.security import generate_password_hash, check_password_hash

USERS_FILE = os.path.join(os.path.dirname(__file__), '../data/users.xlsx')


if not os.path.exists(USERS_FILE):
	df = pd.DataFrame(columns=["username", "password_hash", "role", "batch"])
	df.to_excel(USERS_FILE, index=False)



def add_user(username, email, password, role='teacher'):
	df = pd.read_excel(USERS_FILE)
	if username in df['username'].values or email in df.get('email', pd.Series()).values:
		return False
	password_hash = generate_password_hash(password)
	new_row = {
		"username": username,
		"email": email,
		"password_hash": password_hash,
		"role": role,
		"batch": '' if role == 'teacher' else None
	}
	df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
	df.to_excel(USERS_FILE, index=False)
	return True


def authenticate_user(identifier, password, by_email=False):
	df = pd.read_excel(USERS_FILE)
	if by_email:
		user = df[df['email'] == identifier]
	else:
		user = df[df['username'] == identifier]
	if user.empty:
		return None
	if password is None:
		# Just return user info without checking password
		return {
			"username": user.iloc[0]['username'],
			"role": user.iloc[0].get('role', 'teacher'),
			"batch": user.iloc[0].get('batch', ''),
			"email": user.iloc[0].get('email', '')
		}
	if check_password_hash(user.iloc[0]['password_hash'], password):
		return {
			"username": user.iloc[0]['username'],
			"role": user.iloc[0].get('role', 'teacher'),
			"batch": user.iloc[0].get('batch', ''),
			"email": user.iloc[0].get('email', '')
		}
	return None
def set_user_batch(username, batch):
	df = pd.read_excel(USERS_FILE)
	idx = df.index[df['username'] == username]
	if not idx.empty:
		current = str(df.at[idx[0], 'batch']) if not pd.isna(df.at[idx[0], 'batch']) else ''
		batches = [b.strip() for b in current.split(',') if b.strip()] if current else []
		if batch not in batches:
			batches.append(batch)
			df.at[idx[0], 'batch'] = ','.join(batches)
		df.to_excel(USERS_FILE, index=False)
