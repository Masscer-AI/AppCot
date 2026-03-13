"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { Alert, Button, Container, Paper, PasswordInput, Stack, Text, TextInput, Title } from "@mantine/core";

export default function LoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    setLoading(true);
    setError(null);
    try {
      const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8009";
      const res = await fetch(`${apiBaseUrl}/api/auth/login`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: email.trim(), password }),
      });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { detail?: string };
        throw new Error(data.detail ?? `Error ${res.status}`);
      }
      router.push("/cotizaciones");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado");
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter") handleLogin();
  };

  return (
    <Container size="xs" py={80}>
      <Stack gap="lg">
        <Title order={2} ta="center">
          AppCot
        </Title>
        <Paper withBorder radius="md" p="xl">
          <Stack gap="md">
            <TextInput
              label="Correo"
              placeholder="correo@empresa.com"
              value={email}
              onChange={(e) => setEmail(e.currentTarget.value)}
              onKeyDown={handleKeyDown}
              autoComplete="email"
            />
            <PasswordInput
              label="Contraseña"
              placeholder="••••••••"
              value={password}
              onChange={(e) => setPassword(e.currentTarget.value)}
              onKeyDown={handleKeyDown}
              autoComplete="current-password"
            />
            <Button onClick={handleLogin} loading={loading} fullWidth mt="xs">
              Iniciar sesión
            </Button>
          </Stack>
        </Paper>

        {error ? (
          <Alert color="red">
            <Text size="sm">{error}</Text>
          </Alert>
        ) : null}
      </Stack>
    </Container>
  );
}
