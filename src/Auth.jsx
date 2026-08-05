import React, { useState } from 'react';
import { supabase } from './supabaseClient';

export default function Auth() {
  const [loading, setLoading] = useState(false);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [isSignUp, setIsSignUp] = useState(false);
  const [mensaje, setMensaje] = useState('');

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    setMensaje('');

    if (isSignUp) {
      // REGISTRO
      const { error } = await supabase.auth.signUp({
        email,
        password,
      });
      if (error) setMensaje(error.message);
      else setMensaje('¡Registro exitoso! Revisá tu email si requiere confirmación.');
    } else {
      // INICIO DE SESIÓN
      const { error } = await supabase.auth.signInWithPassword({
        email,
        password,
      });
      if (error) setMensaje(error.message);
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-stone-950 flex items-center justify-center p-4">
      <div className="w-full max-w-md bg-stone-900 border border-stone-800 rounded-xl p-6 shadow-xl text-white">
        <h2 className="text-2xl font-bold text-center mb-6 text-amber-500">
          {isSignUp ? 'Crear Cuenta' : 'Iniciar Sesión'}
        </h2>

        {mensaje && (
          <div className="bg-amber-500/10 border border-amber-500/20 text-amber-400 p-3 rounded-lg text-sm mb-4 text-center">
            {mensaje}
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs uppercase tracking-wider text-stone-400 mb-1">Email</label>
            <input
              type="email"
              required
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full bg-stone-950 border border-stone-800 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
              placeholder="tu@email.com"
            />
          </div>

          <div>
            <label className="block text-xs uppercase tracking-wider text-stone-400 mb-1">Contraseña</label>
            <input
              type="password"
              required
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full bg-stone-950 border border-stone-800 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
              placeholder="••••••••"
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-amber-500 hover:bg-amber-600 text-stone-950 font-bold py-2 px-4 rounded-lg transition-colors disabled:opacity-50"
          >
            {loading ? 'Procesando...' : isSignUp ? 'Registrarse' : 'Ingresar'}
          </button>
        </form>

        <button
          onClick={() => setIsSignUp(!isSignUp)}
          className="w-full text-center text-sm text-stone-400 hover:text-white mt-4 underline"
        >
          {isSignUp ? '¿Ya tenés cuenta? Iniciá sesión' : '¿No tenés cuenta? Registrate acá'}
        </button>
      </div>
    </div>
  );
}
